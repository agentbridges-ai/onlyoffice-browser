import crypto from 'node:crypto';
import { describe, expect, it, vi } from 'vitest';
import {
  EDITOR_SHELL_CACHE_NAME,
  EDITOR_SHELL_MAX_ASSET_BYTES,
  EDITOR_SHELL_MAX_TOTAL_BYTES,
  extractEditorShellDependencies,
  matchEditorShell,
  primeEditorShell,
  releaseIdFromEditorShellPath,
  verifyEditorShell,
} from '../../src/lib/editor-shell-cache';

type StoredResponse = {
  body: ArrayBuffer;
  status: number;
  statusText: string;
  headers: Headers;
};

class MemoryCache {
  readonly responses = new Map<string, StoredResponse>();
  failNextMatch = false;
  failNextPut = false;

  async put(request: RequestInfo | URL, response: Response): Promise<void> {
    if (this.failNextPut) {
      this.failNextPut = false;
      throw new Error('simulated pointer commit failure');
    }
    const body = await response.arrayBuffer();
    this.responses.set(String(request), {
      body,
      status: response.status,
      statusText: response.statusText,
      headers: new Headers(response.headers),
    });
  }

  async match(request: RequestInfo | URL): Promise<Response | undefined> {
    if (this.failNextMatch) {
      this.failNextMatch = false;
      throw new Error('simulated transient cache read failure');
    }
    const stored = this.responses.get(String(request));
    return stored
      ? new Response(stored.body.slice(0), {
          status: stored.status,
          statusText: stored.statusText,
          headers: stored.headers,
        })
      : undefined;
  }

  async delete(request: RequestInfo | URL): Promise<boolean> {
    return this.responses.delete(String(request));
  }
}

class MemoryCacheStorage {
  readonly caches = new Map<string, MemoryCache>();

  async open(name: string): Promise<MemoryCache> {
    let cache = this.caches.get(name);
    if (!cache) {
      cache = new MemoryCache();
      this.caches.set(name, cache);
    }
    return cache;
  }

  async delete(name: string): Promise<boolean> {
    return this.caches.delete(name);
  }

  generation(releaseId: string): MemoryCache {
    const entry = [...this.caches].find(([name]) =>
      name.startsWith(`${EDITOR_SHELL_CACHE_NAME}:generation:${releaseId}:`),
    );
    if (!entry) throw new Error(`Missing generation for ${releaseId}`);
    return entry[1];
  }

  generationNames(releaseId: string): string[] {
    return [...this.caches.keys()].filter((name) =>
      name.startsWith(`${EDITOR_SHELL_CACHE_NAME}:generation:${releaseId}:`),
    );
  }
}

const digest = (value: string | Uint8Array) =>
  crypto
    .createHash('sha256')
    .update(typeof value === 'string' ? Buffer.from(value) : value)
    .digest('hex');

type FixtureAsset = {
  path: string;
  body: string;
  bytes?: number;
  mime?: string;
  sha256?: string;
};

function mimeFor(path: string): string {
  if (path.endsWith('.html')) return 'text/html; charset=utf-8';
  if (path.endsWith('.css')) return 'text/css; charset=utf-8';
  if (path.endsWith('.js')) return 'text/javascript; charset=utf-8';
  return 'application/octet-stream';
}

function defaultAssets(suffix = 'a'): FixtureAsset[] {
  return [
    {
      path: 'office-host.html',
      body: `<script src="./assets/officeHost-${suffix}.js"></script><link href="./assets/base-${suffix}.css">`,
    },
    {
      path: 'editor-shell-prime.html',
      body: `<script src="./assets/editorShellPrime-${suffix}.js"></script>`,
    },
    { path: `assets/base-${suffix}.css`, body: `body{--release:${suffix}}` },
    { path: `assets/editorShellPrime-${suffix}.js`, body: `export const prime="${suffix}"` },
    { path: `assets/officeHost-${suffix}.js`, body: `export const host="${suffix}"` },
  ];
}

function releaseFixture(
  releaseId: string,
  assets = defaultAssets(),
): {
  fetch: ReturnType<typeof vi.fn>;
  assets: Map<string, FixtureAsset>;
  manifestText: string;
} {
  const byPath = new Map(assets.map((asset) => [asset.path, asset]));
  const manifestText = JSON.stringify({
    version: 5,
    releaseId,
    assets: assets.map((asset) => ({
      path: asset.path,
      bytes: asset.bytes ?? Buffer.byteLength(asset.body),
      mime: asset.mime ?? mimeFor(asset.path),
      sha256: asset.sha256 ?? digest(asset.body),
      profile: 'base',
      chunk: 'base-001',
      packageOffset: 0,
      representations: {
        whole: {
          bytes: asset.bytes ?? Buffer.byteLength(asset.body),
          sha256: asset.sha256 ?? digest(asset.body),
        },
      },
    })),
  });
  const fetch = vi.fn(async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    if (url.pathname === `/releases/${releaseId}/manifest.json`) {
      return new Response(manifestText, {
        headers: {
          'content-length': String(Buffer.byteLength(manifestText)),
          'content-type': 'application/json; charset=utf-8',
        },
      });
    }
    const prefix = `/r/${releaseId}/`;
    const path = url.pathname.startsWith(prefix) ? url.pathname.slice(prefix.length) : '';
    const asset = byPath.get(path);
    if (!asset) return new Response('missing', { status: 404 });
    return new Response(asset.body, {
      headers: {
        'content-length': String(Buffer.byteLength(asset.body)),
        'content-type': asset.mime ?? mimeFor(path),
        'x-onlyoffice-asset-version': releaseId,
      },
    });
  });
  return { fetch, assets: byPath, manifestText };
}

describe('editor shell cache', () => {
  it('extracts only deterministic Vite JavaScript and CSS dependencies', () => {
    expect(
      extractEditorShellDependencies(`
        <script src="./assets/officeHost-v1-a.js"></script>
        <link href="./assets/base-v1-b.css">
        <link href="./assets/base-v1-b.css">
      `),
    ).toEqual(['assets/base-v1-b.css', 'assets/officeHost-v1-a.js']);
    expect(() => extractEditorShellDependencies('<script src="./assets/../secret.js"></script>')).toThrow(/unsafe/);
    expect(() => extractEditorShellDependencies('<link href="./assets/font.woff2">')).toThrow(/unsafe/);
  });

  it('parses only exact release-bound host and prime paths', () => {
    expect(releaseIdFromEditorShellPath('/r/v0.6.0%2B1/office-host.html', 'office-host.html')).toBe('v0.6.0+1');
    expect(releaseIdFromEditorShellPath('/r/../../office-host.html', 'office-host.html')).toBeNull();
    expect(releaseIdFromEditorShellPath('/r/release/other.html', 'office-host.html')).toBeNull();
  });

  it('stages and deeply verifies a manifest-bound shell before atomically publishing it', async () => {
    const storage = new MemoryCacheStorage();
    const fixture = releaseFixture('release-a');
    const result = await primeEditorShell({
      releaseId: 'release-a',
      origin: 'https://aries.getpi.work',
      manifestOrigin: 'https://onlyoffice.getpi.work',
      cacheStorage: storage as unknown as CacheStorage,
      fetch: fixture.fetch as typeof fetch,
      createGenerationId: () => '1'.repeat(32),
    });
    expect(result.cachedPaths).toEqual([
      'assets/base-a.css',
      'assets/editorShellPrime-a.js',
      'assets/officeHost-a.js',
      'editor-shell-prime.html',
      'office-host.html',
    ]);
    expect(result.cachedBytes).toBe([...fixture.assets.values()].reduce((sum, asset) => sum + asset.body.length, 0));
    expect(fixture.fetch).toHaveBeenCalledTimes(6);
    expect(String(fixture.fetch.mock.calls[0][0])).toBe(
      'https://onlyoffice.getpi.work/releases/release-a/manifest.json',
    );
    expect(
      fixture.fetch.mock.calls.slice(1).every(([input]) => String(input).startsWith('https://aries.getpi.work/')),
    ).toBe(true);
    expect(storage.generationNames('release-a')).toEqual([
      `${EDITOR_SHELL_CACHE_NAME}:generation:release-a:${'1'.repeat(32)}`,
    ]);
    const matchedHost = await matchEditorShell(
      new Request('https://aries.getpi.work/r/release-a/assets/officeHost-a.js'),
      'release-a',
      storage as unknown as CacheStorage,
    );
    expect(matchedHost?.headers.get('x-onlyoffice-editor-shell-cache')).toBe('1');
    expect(await matchedHost?.text()).toBe('export const host="a"');
    await expect(
      verifyEditorShell({
        releaseId: 'release-a',
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
      }),
    ).resolves.toEqual(result);
  });

  it('rematerializes a persisted shell response with actual decoded length', async () => {
    const storage = new MemoryCacheStorage();
    const fixture = releaseFixture('release-a');
    await primeEditorShell({
      releaseId: 'release-a',
      origin: 'https://aries.getpi.work',
      cacheStorage: storage as unknown as CacheStorage,
      fetch: fixture.fetch as typeof fetch,
      createGenerationId: () => '1'.repeat(32),
    });
    const url = 'https://aries.getpi.work/r/release-a/office-host.html';
    const stored = storage.generation('release-a').responses.get(url)!;
    stored.headers.set('content-encoding', 'br');
    stored.headers.set('content-length', '1');
    stored.headers.set('content-security-policy', "default-src 'none'");

    const response = await matchEditorShell(new Request(url), 'release-a', storage as unknown as CacheStorage);
    expect(response?.headers.get('content-encoding')).toBeNull();
    expect(response?.headers.get('content-length')).toBe(String(stored.body.byteLength));
    expect(response?.headers.get('content-security-policy')).toBe("default-src 'none'");
    await expect(response?.text()).resolves.toContain('officeHost-a.js');
  });

  it('serves a release-bound Host navigation whose client identity is carried in the URL fragment', async () => {
    const storage = new MemoryCacheStorage();
    const fixture = releaseFixture('release-a');
    await primeEditorShell({
      releaseId: 'release-a',
      origin: 'https://aries.getpi.work',
      cacheStorage: storage as unknown as CacheStorage,
      fetch: fixture.fetch as typeof fetch,
      createGenerationId: () => '1'.repeat(32),
    });

    const response = await matchEditorShell(
      new Request(
        'https://aries.getpi.work/r/release-a/office-host.html#sessionId=office-editor-1&parentOrigin=https%3A%2F%2Fonlyoffice.getpi.work',
      ),
      'release-a',
      storage as unknown as CacheStorage,
    );
    expect(response?.status).toBe(200);
    expect(response?.headers.get('x-onlyoffice-editor-shell-cache')).toBe('1');
    await expect(response?.text()).resolves.toContain('officeHost-a.js');
  });

  it('accepts decoded immutable responses without Content-Length and still verifies bytes and SHA-256', async () => {
    const storage = new MemoryCacheStorage();
    const fixture = releaseFixture('release-decoded');
    fixture.fetch.mockImplementation(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (url.pathname === '/releases/release-decoded/manifest.json') {
        return new Response(fixture.manifestText, {
          headers: { 'content-type': 'application/json; charset=utf-8' },
        });
      }
      const path = url.pathname.replace('/r/release-decoded/', '');
      const asset = fixture.assets.get(path);
      if (!asset) return new Response('missing', { status: 404 });
      return new Response(asset.body, {
        headers: {
          'content-type': asset.mime ?? mimeFor(path),
          'x-onlyoffice-asset-version': 'release-decoded',
        },
      });
    });

    await expect(
      primeEditorShell({
        releaseId: 'release-decoded',
        origin: 'https://aries.getpi.work',
        manifestOrigin: 'https://onlyoffice.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
        fetch: fixture.fetch as typeof fetch,
      }),
    ).resolves.toMatchObject({ releaseId: 'release-decoded' });
  });

  it('keeps a verified same-release active generation when a replacement commit fails', async () => {
    const storage = new MemoryCacheStorage();
    const fixture = releaseFixture('release-a');
    const options = {
      releaseId: 'release-a',
      origin: 'https://aries.getpi.work',
      cacheStorage: storage as unknown as CacheStorage,
      fetch: fixture.fetch as typeof fetch,
    };
    await primeEditorShell({ ...options, createGenerationId: () => '1'.repeat(32) });
    const activeGeneration = storage.generation('release-a');
    activeGeneration.failNextMatch = true;
    const metadata = await storage.open(EDITOR_SHELL_CACHE_NAME);
    metadata.failNextPut = true;

    await expect(primeEditorShell({ ...options, createGenerationId: () => '2'.repeat(32) })).rejects.toThrow(
      /commit failure/,
    );
    expect(storage.generationNames('release-a')).toEqual([
      `${EDITOR_SHELL_CACHE_NAME}:generation:release-a:${'1'.repeat(32)}`,
    ]);
    await expect(verifyEditorShell(options)).resolves.toMatchObject({
      releaseId: 'release-a',
    });
  });

  it('does not expose a partially staged B and retains A when a B dependency is interrupted', async () => {
    const storage = new MemoryCacheStorage();
    const fixtureA = releaseFixture('release-a', defaultAssets('a'));
    await primeEditorShell({
      releaseId: 'release-a',
      origin: 'https://aries.getpi.work',
      cacheStorage: storage as unknown as CacheStorage,
      fetch: fixtureA.fetch as typeof fetch,
      createGenerationId: () => '1'.repeat(32),
    });

    const fixtureB = releaseFixture('release-b', defaultAssets('b'));
    let releaseBlocked!: () => void;
    const blocked = new Promise<void>((resolve) => {
      releaseBlocked = resolve;
    });
    let dependencyRequested!: () => void;
    const requested = new Promise<void>((resolve) => {
      dependencyRequested = resolve;
    });
    const interruptedFetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
      const url = new URL(String(input));
      if (url.pathname.endsWith('/assets/officeHost-b.js')) {
        dependencyRequested();
        await blocked;
        return new Response('interrupted', { status: 503 });
      }
      return (fixtureB.fetch as unknown as typeof fetch)(input, init);
    };
    const preparingB = primeEditorShell({
      releaseId: 'release-b',
      origin: 'https://aries.getpi.work',
      cacheStorage: storage as unknown as CacheStorage,
      fetch: interruptedFetch as typeof fetch,
      createGenerationId: () => '2'.repeat(32),
    });
    await requested;

    await expect(
      matchEditorShell(
        new Request('https://aries.getpi.work/r/release-b/office-host.html'),
        'release-b',
        storage as unknown as CacheStorage,
      ),
    ).resolves.toBeUndefined();
    await expect(
      matchEditorShell(
        new Request('https://aries.getpi.work/r/release-a/office-host.html'),
        'release-a',
        storage as unknown as CacheStorage,
      ),
    ).resolves.toBeDefined();

    releaseBlocked();
    await expect(preparingB).rejects.toThrow(/failed/);
    expect(storage.generationNames('release-b')).toEqual([]);
    await expect(
      verifyEditorShell({
        releaseId: 'release-a',
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
      }),
    ).resolves.toMatchObject({ releaseId: 'release-a' });
  });

  it('keeps A active while a complete immutable B generation is prepared', async () => {
    const storage = new MemoryCacheStorage();
    for (const [releaseId, suffix, generation] of [
      ['release-a', 'a', '1'.repeat(32)],
      ['release-b', 'b', '2'.repeat(32)],
    ]) {
      const fixture = releaseFixture(releaseId, defaultAssets(suffix));
      await primeEditorShell({
        releaseId,
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
        fetch: fixture.fetch as typeof fetch,
        createGenerationId: () => generation,
      });
    }
    await expect(
      verifyEditorShell({
        releaseId: 'release-a',
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
      }),
    ).resolves.toMatchObject({ releaseId: 'release-a' });
    await expect(
      verifyEditorShell({
        releaseId: 'release-b',
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
      }),
    ).resolves.toMatchObject({ releaseId: 'release-b' });
  });

  it('rejects a manifest dependency above the per-asset byte limit before downloading it', async () => {
    const storage = new MemoryCacheStorage();
    const assets = defaultAssets();
    const dependency = assets.find((asset) => asset.path === 'assets/officeHost-a.js')!;
    dependency.bytes = EDITOR_SHELL_MAX_ASSET_BYTES + 1;
    dependency.sha256 = digest(dependency.body);
    const fixture = releaseFixture('release-a', assets);
    await expect(
      primeEditorShell({
        releaseId: 'release-a',
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
        fetch: fixture.fetch as typeof fetch,
        createGenerationId: () => '1'.repeat(32),
      }),
    ).rejects.toThrow(/invalid editor shell asset/);
    expect(fixture.fetch.mock.calls.some(([input]) => String(input).includes('/assets/officeHost-a.js'))).toBe(false);
  });

  it('rejects a manifest dependency graph above the total byte limit before downloading dependencies', async () => {
    const dependencyPaths = Array.from({ length: 5 }, (_, index) => `assets/large-${index}.js`);
    const host = dependencyPaths.map((path) => `<script src="./${path}"></script>`).join('');
    const assets: FixtureAsset[] = [
      { path: 'office-host.html', body: host },
      { path: 'editor-shell-prime.html', body: '<script src="./assets/large-0.js"></script>' },
      ...dependencyPaths.map((path) => ({
        path,
        body: 'x',
        bytes: Math.floor(EDITOR_SHELL_MAX_TOTAL_BYTES / 4),
        sha256: digest('x'),
      })),
    ];
    const storage = new MemoryCacheStorage();
    const fixture = releaseFixture('release-a', assets);
    await expect(
      primeEditorShell({
        releaseId: 'release-a',
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
        fetch: fixture.fetch as typeof fetch,
        createGenerationId: () => '1'.repeat(32),
      }),
    ).rejects.toThrow(/bounded total size/);
    expect(fixture.fetch.mock.calls.some(([input]) => String(input).includes('/assets/large-'))).toBe(false);
  });

  it.each([
    {
      label: 'MIME',
      mutate(assets: FixtureAsset[]) {
        assets.find((asset) => asset.path === 'assets/officeHost-a.js')!.mime = 'application/octet-stream';
      },
      expected: /invalid editor shell asset/,
    },
    {
      label: 'digest',
      mutate(assets: FixtureAsset[]) {
        assets.find((asset) => asset.path === 'assets/officeHost-a.js')!.sha256 = '0'.repeat(64);
      },
      expected: /SHA-256/,
    },
    {
      label: 'path traversal',
      mutate(assets: FixtureAsset[]) {
        assets.push({ path: 'assets/../secret.js', body: 'secret' });
      },
      expected: /invalid asset path/,
    },
  ])('rejects a manifest or response $label anomaly', async ({ mutate, expected }) => {
    const assets = defaultAssets();
    mutate(assets);
    const fixture = releaseFixture('release-a', assets);
    await expect(
      primeEditorShell({
        releaseId: 'release-a',
        origin: 'https://aries.getpi.work',
        cacheStorage: new MemoryCacheStorage() as unknown as CacheStorage,
        fetch: fixture.fetch as typeof fetch,
        createGenerationId: () => '1'.repeat(32),
      }),
    ).rejects.toThrow(expected);
  });

  it('detects cached body tampering even when the headers and byte count still look valid', async () => {
    const storage = new MemoryCacheStorage();
    const fixture = releaseFixture('release-a');
    await primeEditorShell({
      releaseId: 'release-a',
      origin: 'https://aries.getpi.work',
      cacheStorage: storage as unknown as CacheStorage,
      fetch: fixture.fetch as typeof fetch,
      createGenerationId: () => '1'.repeat(32),
    });
    const generation = storage.generation('release-a');
    const url = 'https://aries.getpi.work/r/release-a/assets/officeHost-a.js';
    const stored = generation.responses.get(url)!;
    stored.body = new TextEncoder().encode('x'.repeat(stored.body.byteLength)).buffer;

    await expect(
      verifyEditorShell({
        releaseId: 'release-a',
        origin: 'https://aries.getpi.work',
        cacheStorage: storage as unknown as CacheStorage,
      }),
    ).rejects.toThrow(/SHA-256/);
  });
});
