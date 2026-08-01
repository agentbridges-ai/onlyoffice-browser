import { describe, expect, it, vi } from 'vitest';
import { AssetSpanReader, type AssetContentObjectSource } from '../../src/lib/asset-span-reader';
import type { AssetContentMapping } from '../../src/lib/release-content-model';

const digestA = 'a'.repeat(64);
const digestB = 'b'.repeat(64);

const mapping: AssetContentMapping = {
  releaseId: 'release-a',
  path: 'wasm/runtime.wasm',
  assetSha256: 'c'.repeat(64),
  assetBytes: 10,
  representationId: 'package',
  representationKind: 'package',
  spans: [
    { objectSha256: digestA, objectOffset: 2, assetOffset: 0, bytes: 4 },
    { objectSha256: digestB, objectOffset: 1, assetOffset: 4, bytes: 6 },
  ],
};

function stream(chunks: number[][], cancelled?: (reason: unknown) => void): ReadableStream<Uint8Array> {
  return new ReadableStream({
    type: 'bytes',
    start(controller) {
      for (const chunk of chunks) controller.enqueue(Uint8Array.from(chunk));
      controller.close();
    },
    cancel(reason) {
      cancelled?.(reason);
    },
  });
}

describe('AssetSpanReader', () => {
  it('streams a bounded range across object spans without assembling the asset', async () => {
    const loader = vi.fn(async (sha256: string): Promise<AssetContentObjectSource> => {
      if (sha256 === digestA) {
        return {
          bytes: 8,
          stream: stream([
            [90, 91, 1],
            [2, 3, 4, 92, 93],
          ]),
        };
      }
      return {
        bytes: 8,
        stream: stream([
          [80, 5, 6],
          [7, 8, 9, 10, 81],
        ]),
      };
    });
    const reader = new AssetSpanReader(mapping, { start: 2, end: 8 }, loader);
    const pieces: number[][] = [];
    while (true) {
      const piece = await reader.read(3);
      if (!piece) break;
      expect(piece.byteLength).toBeLessThanOrEqual(3);
      pieces.push([...piece]);
    }
    expect(pieces.flat()).toEqual([3, 4, 5, 6, 7, 8, 9]);
    expect(reader.bytesRead).toBe(7);
    expect(loader).toHaveBeenCalledTimes(2);
  });

  it('cancels the active object reader immediately', async () => {
    const cancelled = vi.fn();
    const reader = new AssetSpanReader(mapping, { start: 0, end: 9 }, async () => ({
      bytes: 20,
      stream: new ReadableStream({
        type: 'bytes',
        start(controller) {
          controller.enqueue(Uint8Array.from([0, 0, 1, 2, 3, 4, 5, 6]));
        },
        cancel(reason) {
          cancelled(reason);
        },
      }),
    }));
    expect(await reader.read(1)).toEqual(Uint8Array.from([1]));
    await reader.cancel('consumer');
    expect(cancelled).toHaveBeenCalledWith('consumer');
    expect(await reader.read(1)).toBeNull();
  });

  it('fails when a stored object is truncated or its ledger size cannot cover the span', async () => {
    const truncated = new AssetSpanReader(mapping, { start: 0, end: 3 }, async () => ({
      bytes: 8,
      stream: stream([[0, 0, 1]]),
    }));
    expect(await truncated.read(4)).toEqual(Uint8Array.from([1]));
    await expect(truncated.read(4)).rejects.toThrow(/ended before/);

    const wrongLedger = new AssetSpanReader(mapping, { start: 0, end: 3 }, async () => ({
      bytes: 5,
      stream: stream([[0, 0, 1, 2, 3, 4]]),
    }));
    await expect(wrongLedger.read(4)).rejects.toThrow(/does not cover/);
  });

  it('seeks to a late range in logarithmic span-index reads', async () => {
    const baseSpans = Array.from({ length: 1_024 }, (_, index) => ({
      objectSha256: index.toString(16).padStart(64, '0'),
      objectOffset: 0,
      assetOffset: index,
      bytes: 1,
    }));
    let indexedReads = 0;
    const spans = new Proxy(baseSpans, {
      get(target, property, receiver) {
        if (typeof property === 'string' && /^(?:0|[1-9][0-9]*)$/.test(property)) indexedReads += 1;
        return Reflect.get(target, property, receiver);
      },
    });
    const largeMapping: AssetContentMapping = {
      releaseId: 'release-large',
      path: 'wasm/large.bin',
      assetSha256: 'f'.repeat(64),
      assetBytes: spans.length,
      representationId: 'fastcdc',
      representationKind: 'fastcdc',
      spans,
    };
    const loader = vi.fn(
      async (): Promise<AssetContentObjectSource> => ({
        bytes: 1,
        stream: stream([[42]]),
      }),
    );

    const reader = new AssetSpanReader(largeMapping, { start: 1_023, end: 1_023 }, loader);
    expect(await reader.read(1)).toEqual(Uint8Array.of(42));
    expect(await reader.read(1)).toBeNull();

    expect(loader).toHaveBeenCalledTimes(1);
    expect(indexedReads).toBeLessThanOrEqual(12);
  });

  it('uses BYOB demand so a large stored object never becomes an eager whole-object buffer', async () => {
    const objectBytes = 32 * 1024 * 1024;
    const windowBytes = 1024 * 1024;
    let producedBytes = 0;
    let largestRequestedView = 0;
    const cancelled = vi.fn();
    const largeMapping: AssetContentMapping = {
      releaseId: 'release-large-object',
      path: 'wasm/large-object.bin',
      assetSha256: digestA,
      assetBytes: objectBytes,
      representationId: 'whole',
      representationKind: 'whole',
      spans: [{ objectSha256: digestA, objectOffset: 0, assetOffset: 0, bytes: objectBytes }],
    };
    const source = new ReadableStream<Uint8Array>({
      type: 'bytes',
      pull(controller) {
        const request = (controller as ReadableByteStreamController).byobRequest;
        if (!request?.view) throw new Error('expected a BYOB request');
        largestRequestedView = Math.max(largestRequestedView, request.view.byteLength);
        const count = Math.min(request.view.byteLength, objectBytes - producedBytes);
        const view = new Uint8Array(request.view.buffer, request.view.byteOffset, count);
        view.fill(42);
        producedBytes += count;
        request.respond(count);
        if (producedBytes === objectBytes) controller.close();
      },
      cancel(reason) {
        cancelled(reason);
      },
    });
    const reader = new AssetSpanReader(largeMapping, { start: 0, end: objectBytes - 1 }, async () => ({
      bytes: objectBytes,
      stream: source,
    }));

    expect((await reader.read(windowBytes))?.byteLength).toBe(windowBytes);
    expect((await reader.read(windowBytes))?.byteLength).toBe(windowBytes);
    expect(producedBytes).toBe(2 * windowBytes);
    expect(largestRequestedView).toBe(windowBytes);

    await reader.cancel('stop-before-whole-object');
    expect(cancelled).toHaveBeenCalledWith('stop-before-whole-object');
    expect(producedBytes).toBeLessThan(objectBytes);
  });

  it('fails closed before pulling from a non-byte stream that could emit an oversized chunk', async () => {
    const oversizedBytes = 32 * 1024 * 1024;
    const pulled = vi.fn();
    const cancelled = vi.fn();
    const unsafeStream = new ReadableStream<Uint8Array>(
      {
        pull(controller) {
          pulled();
          controller.enqueue(new Uint8Array(oversizedBytes));
        },
        cancel(reason) {
          cancelled(reason);
        },
      },
      { highWaterMark: 0 },
    );
    const reader = new AssetSpanReader(mapping, { start: 0, end: 3 }, async () => ({
      bytes: oversizedBytes,
      stream: unsafeStream,
    }));

    await expect(reader.read(1024 * 1024)).rejects.toThrow(/bounded BYOB reads/);
    expect(pulled).not.toHaveBeenCalled();
    expect(cancelled).toHaveBeenCalledOnce();
  });
});
