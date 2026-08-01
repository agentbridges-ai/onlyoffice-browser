import { afterEach, describe, expect, it, vi } from 'vitest';
import { EditorOriginPrimeError, EditorOriginPrimer, editorOriginForSlot } from '../../src/lib/editor-origin-primer';
import { EDITOR_SHELL_MAX_TOTAL_BYTES } from '../../src/lib/editor-shell-cache';
import { RESOURCE_BROKER_PROTOCOL } from '../../src/lib/resource-broker-protocol';

type MessageListener = (event: MessageEvent) => void;

function primerHarness(
  options: {
    brokerReady?: boolean;
    occupied?: boolean;
    cachedPaths?: string[];
    cachedBytes?: number;
    delayMs?: number;
    failure?: { code: 'storage' | 'timeout' | 'cancelled'; stage: 'service-worker' | 'shell-cache' | 'broker-probe' };
  } = {},
) {
  const listeners = new Set<MessageListener>();
  const frames: Array<{ src: string; contentWindow: object; removed: boolean }> = [];
  const browserWindow = {
    location: { origin: 'https://onlyoffice.getpi.work' },
    setTimeout: window.setTimeout.bind(window),
    clearTimeout: window.clearTimeout.bind(window),
    addEventListener(type: string, listener: MessageListener) {
      if (type === 'message') listeners.add(listener);
    },
    removeEventListener(type: string, listener: MessageListener) {
      if (type === 'message') listeners.delete(listener);
    },
  } as unknown as Window;
  const browserDocument = {
    createElement() {
      const frame = {
        src: '',
        contentWindow: {},
        hidden: false,
        tabIndex: 0,
        removed: false,
        setAttribute() {},
        remove() {
          frame.removed = true;
        },
      };
      frames.push(frame);
      return frame;
    },
    body: {
      append(frame: (typeof frames)[number]) {
        const url = new URL(frame.src);
        const fragment = new URLSearchParams(url.hash.slice(1));
        const respond = () => {
          const failure = options.failure;
          const event = {
            source: frame.contentWindow,
            origin: url.origin,
            data: failure
              ? {
                  protocol: RESOURCE_BROKER_PROTOCOL,
                  type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_FAILED',
                  origin: url.origin,
                  releaseId: fragment.get('releaseId'),
                  sessionId: fragment.get('sessionId'),
                  code: failure.code,
                  stage: failure.stage,
                }
              : {
                  protocol: RESOURCE_BROKER_PROTOCOL,
                  type: 'ONLYOFFICE_EDITOR_SHELL_PRIMED',
                  origin: url.origin,
                  releaseId: fragment.get('releaseId'),
                  sessionId: fragment.get('sessionId'),
                  brokerReady: options.brokerReady ?? true,
                  occupied: options.occupied ?? false,
                  serviceWorkerVersion: 'sw-test',
                  cachedPaths: options.cachedPaths ?? [
                    'office-host.html',
                    'editor-shell-prime.html',
                    'assets/officeHost.js',
                  ],
                  cachedBytes: options.cachedBytes ?? 42,
                },
          } as unknown as MessageEvent;
          for (const listener of listeners) listener(event);
        };
        if (options.delayMs === undefined) queueMicrotask(respond);
        else window.setTimeout(respond, options.delayMs);
      },
    },
  } as unknown as Document;
  return { browserWindow, browserDocument, frames };
}

afterEach(() => {
  vi.useRealTimers();
});

describe('EditorOriginPrimer', () => {
  it('uses the fixed production constellation origins', () => {
    expect(editorOriginForSlot('aries', 'https://onlyoffice.getpi.work')).toBe('https://aries.getpi.work');
    expect(editorOriginForSlot('gemini', 'https://onlyoffice.getpi.work')).toBe('https://gemini.getpi.work');
  });

  it('uses the same local protocol and port for isolated editor origins', () => {
    expect(editorOriginForSlot('taurus', 'http://onlyoffice.localhost:8787', true)).toBe(
      'http://host-taurus.office.localhost:8787',
    );
    expect(() => editorOriginForSlot('aries', 'http://evil.localhost:8787', true)).toThrow(/Invalid/);
  });

  it('reports an occupied origin without replacing its pinned Broker connection', async () => {
    const harness = primerHarness({ brokerReady: false, occupied: true });
    const primer = new EditorOriginPrimer({
      window: harness.browserWindow,
      document: harness.browserDocument,
      randomUUID: () => 'session-a',
    });

    await expect(primer.primeRelease('release-b', ['aries'])).resolves.toEqual([
      expect.objectContaining({
        releaseId: 'release-b',
        brokerReady: false,
        occupied: true,
      }),
    ]);
  });

  it('requires the exact per-origin session and a successful Broker probe before accepting prime results', async () => {
    const harness = primerHarness();
    let id = 0;
    const primer = new EditorOriginPrimer({
      window: harness.browserWindow,
      document: harness.browserDocument,
      randomUUID: () => `session-${++id}`,
    });
    const results = await primer.primeRelease('release-a', ['aries', 'taurus', 'gemini']);

    expect(results).toHaveLength(3);
    expect(results).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          origin: 'https://aries.getpi.work',
          releaseId: 'release-a',
          brokerReady: true,
        }),
      ]),
    );
    expect(harness.frames.every((frame) => frame.removed)).toBe(true);
    expect(
      harness.frames.every((frame) => {
        const url = new URL(frame.src);
        const fragment = new URLSearchParams(url.hash.slice(1));
        return (
          url.search === '?install=1' &&
          fragment.get('mode') === 'install' &&
          fragment.get('releaseId') === 'release-a' &&
          fragment.get('sessionId')?.startsWith('prime-session-')
        );
      }),
    ).toBe(true);
  });

  it('uses the cached versioned prime page without the install query when verifying after restart', async () => {
    const harness = primerHarness();
    const primer = new EditorOriginPrimer({
      window: harness.browserWindow,
      document: harness.browserDocument,
      randomUUID: () => 'session-a',
    });
    await primer.verifyRelease('release-a', ['aries']);

    const url = new URL(harness.frames[0].src);
    const fragment = new URLSearchParams(url.hash.slice(1));
    expect(url.search).toBe('');
    expect(fragment.get('mode')).toBe('verify');
  });

  it('reports progress after every origin so four 10-second batches remain active beyond 30 seconds', async () => {
    vi.useFakeTimers();
    const harness = primerHarness({ delayMs: 10_000 });
    let id = 0;
    const primer = new EditorOriginPrimer({
      window: harness.browserWindow,
      document: harness.browserDocument,
      randomUUID: () => `session-${++id}`,
    });
    const onProgress = vi.fn();
    const priming = primer.primeRelease('release-a', undefined, onProgress);

    for (let batch = 1; batch <= 4; batch += 1) {
      await vi.advanceTimersByTimeAsync(10_000);
      expect(onProgress).toHaveBeenCalledTimes(batch * 3);
    }

    await expect(priming).resolves.toHaveLength(12);
    expect(harness.frames.every((frame) => frame.removed)).toBe(true);
  });

  it('rejects a shell response that did not complete the actual Broker probe', async () => {
    const harness = primerHarness({ brokerReady: false });
    const primer = new EditorOriginPrimer({
      window: harness.browserWindow,
      document: harness.browserDocument,
      randomUUID: () => 'session-a',
    });
    await expect(primer.primeRelease('release-a', ['aries'])).rejects.toThrow(/Invalid/);
  });

  it('fails immediately with the structured editor-origin stage instead of waiting for timeout', async () => {
    const harness = primerHarness({ failure: { code: 'cancelled', stage: 'broker-probe' } });
    const primer = new EditorOriginPrimer({
      window: harness.browserWindow,
      document: harness.browserDocument,
      randomUUID: () => 'session-a',
    });
    const failure = primer.primeRelease('release-a', ['aries']);
    await expect(failure).rejects.toThrow('Editor origin prime failed (broker-probe/cancelled)');
    await expect(failure).rejects.toMatchObject({
      name: 'EditorOriginPrimeError',
      code: 'cancelled',
      stage: 'broker-probe',
      origin: 'https://aries.getpi.work',
    } satisfies Partial<EditorOriginPrimeError>);
  });

  it('rejects an unbounded or unsafe shell-cache result before it reaches readiness', async () => {
    const unsafe = primerHarness({
      cachedPaths: ['office-host.html', 'editor-shell-prime.html', 'assets/../secret.js'],
    });
    const primer = new EditorOriginPrimer({
      window: unsafe.browserWindow,
      document: unsafe.browserDocument,
      randomUUID: () => 'session-a',
    });
    await expect(primer.primeRelease('release-a', ['aries'])).rejects.toThrow(/Invalid/);

    const unbounded = primerHarness({ cachedBytes: EDITOR_SHELL_MAX_TOTAL_BYTES + 1 });
    const unboundedPrimer = new EditorOriginPrimer({
      window: unbounded.browserWindow,
      document: unbounded.browserDocument,
      randomUUID: () => 'session-b',
    });
    await expect(unboundedPrimer.primeRelease('release-a', ['aries'])).rejects.toThrow(/Invalid/);
  });
});
