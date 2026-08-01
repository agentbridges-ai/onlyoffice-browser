import type { AssetContentMapping, AssetContentSpan } from './release-content-model';

export interface AssetContentObjectSource {
  bytes: number;
  stream: ReadableStream<Uint8Array>;
}

export type AssetContentObjectLoader = (sha256: string) => Promise<AssetContentObjectSource>;

export interface AssetReadBounds {
  start: number;
  end: number;
}

type ActiveSpan = {
  span: AssetContentSpan;
  reader: ReadableStreamBYOBReader;
  skipBytes: number;
  remainingBytes: number;
};

function assertBounds(mapping: AssetContentMapping, bounds: AssetReadBounds): void {
  if (
    !Number.isSafeInteger(bounds.start) ||
    !Number.isSafeInteger(bounds.end) ||
    bounds.start < 0 ||
    bounds.end < bounds.start ||
    bounds.end >= mapping.assetBytes
  ) {
    throw new RangeError('Asset read bounds are outside the mapped asset');
  }
}

/**
 * Reads an asset directly from its content-addressed spans. It never assembles
 * the complete asset and returns at most one caller-controlled window per
 * `read()` call, so a Service Worker can couple it to consumer backpressure.
 */
export class AssetSpanReader {
  private spanIndex = 0;
  private active: ActiveSpan | null = null;
  private cancelled = false;
  private cancelReason: unknown;
  private finished = false;
  private readBytes = 0;

  constructor(
    private readonly mapping: AssetContentMapping,
    private readonly bounds: AssetReadBounds,
    private readonly loadObject: AssetContentObjectLoader,
  ) {
    assertBounds(mapping, bounds);
    let low = 0;
    let high = mapping.spans.length;
    while (low < high) {
      const middle = low + Math.floor((high - low) / 2);
      const span = mapping.spans[middle];
      if (span.assetOffset + span.bytes <= bounds.start) low = middle + 1;
      else high = middle;
    }
    this.spanIndex = low;
  }

  get bytesRead(): number {
    return this.readBytes;
  }

  async read(maxBytes: number): Promise<Uint8Array | null> {
    if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) {
      throw new RangeError('Read window must be a positive safe integer');
    }
    if (this.cancelled || this.finished) return null;

    while (true) {
      const active = await this.ensureActiveSpan();
      if (!active) {
        this.finished = true;
        return null;
      }

      if (active.skipBytes > 0) {
        const skipped = await this.readBounded(active, Math.min(maxBytes, active.skipBytes));
        if (!skipped) {
          throw new Error(`Content object ${active.span.objectSha256} ended before its mapped span`);
        }
        active.skipBytes -= skipped.byteLength;
        continue;
      }
      if (active.remainingBytes === 0) {
        await this.closeActive();
        continue;
      }

      const output = await this.readBounded(active, Math.min(maxBytes, active.remainingBytes));
      if (!output) {
        throw new Error(`Content object ${active.span.objectSha256} ended before its mapped span`);
      }
      active.remainingBytes -= output.byteLength;
      this.readBytes += output.byteLength;
      if (active.remainingBytes === 0) {
        await this.closeActive();
      }
      return output;
    }
  }

  async cancel(reason?: unknown): Promise<void> {
    if (this.cancelled) return;
    this.cancelled = true;
    this.cancelReason = reason;
    if (this.active) {
      await this.active.reader.cancel(reason);
      this.active = null;
    }
  }

  private async ensureActiveSpan(): Promise<ActiveSpan | null> {
    if (this.active) return this.active;
    while (this.spanIndex < this.mapping.spans.length) {
      const span = this.mapping.spans[this.spanIndex++];
      const spanStart = span.assetOffset;
      const spanEnd = span.assetOffset + span.bytes - 1;
      const overlapStart = Math.max(this.bounds.start, spanStart);
      const overlapEnd = Math.min(this.bounds.end, spanEnd);
      if (overlapEnd < overlapStart) continue;

      const source = await this.loadObject(span.objectSha256);
      if (!Number.isSafeInteger(source.bytes) || source.bytes < 0 || span.objectOffset + span.bytes > source.bytes) {
        throw new Error(`Content object ${span.objectSha256} does not cover its mapped span`);
      }
      if (this.cancelled) {
        await source.stream.cancel(this.cancelReason).catch(() => undefined);
        return null;
      }
      let reader: ReadableStreamBYOBReader;
      try {
        // Fetch and Cache Storage response bodies are readable byte streams.
        // BYOB is required here: a default reader may hand us an arbitrarily
        // large producer chunk which cannot be split without retaining its
        // remainder outside the caller's explicit memory window.
        reader = source.stream.getReader({ mode: 'byob' });
      } catch (error) {
        await source.stream.cancel(error).catch(() => undefined);
        throw new TypeError('Content object stream must support bounded BYOB reads', { cause: error });
      }
      this.active = {
        span,
        reader,
        skipBytes: span.objectOffset + (overlapStart - spanStart),
        remainingBytes: overlapEnd - overlapStart + 1,
      };
      return this.active;
    }
    return null;
  }

  private async closeActive(): Promise<void> {
    if (!this.active) return;
    const active = this.active;
    this.active = null;
    await active.reader.cancel();
  }

  private async readBounded(active: ActiveSpan, maxBytes: number): Promise<Uint8Array | null> {
    const target = new Uint8Array(maxBytes);
    const next = await active.reader.read(target);
    if (next.done) return null;
    const value = next.value;
    if (!(value instanceof Uint8Array) || value.byteLength === 0 || value.byteLength > maxBytes) {
      throw new Error(`Content object ${active.span.objectSha256} violated the bounded byte-stream contract`);
    }
    if (value.byteOffset === 0 && value.buffer instanceof ArrayBuffer && value.buffer.byteLength === value.byteLength) {
      return value;
    }
    return value.slice();
  }
}
