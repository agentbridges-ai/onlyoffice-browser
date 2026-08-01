import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { OfficeRuntimeResourceManager, OfficeRuntimeResourceSnapshot } from '../../src/lib/runtime-resources';
import { officeCopy } from '../../src/pwa/i18n';
import { OfficeResourcePanel } from '../../src/pwa/resource-panel';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const snapshot: OfficeRuntimeResourceSnapshot = {
  packageVersion: '0.5.0',
  assetVersion: 'v0.4.7-test',
  readiness: 'ready',
  packs: [
    { id: 'core', ready: true, completedBytes: 4, totalBytes: 4 },
    { id: 'word', ready: true, completedBytes: 4, totalBytes: 4 },
    { id: 'cell', ready: false, completedBytes: 0, totalBytes: 4 },
    { id: 'slide', ready: false, completedBytes: 0, totalBytes: 4 },
    { id: 'fonts', ready: true, completedBytes: 4, totalBytes: 4 },
  ],
  progress: {
    phase: 'ready',
    completedFiles: 3,
    totalFiles: 5,
    completedBytes: 12,
    totalBytes: 20,
    failedFiles: 0,
    categories: [],
  },
  fonts: [],
  verifiedFontPaths: [],
  operation: null,
  error: null,
  installedRelease: 'v0.4.7-test',
  targetRelease: 'v0.4.7-test',
  availableRelease: 'v0.4.7-test',
  availablePackageVersion: '0.5.0',
  storageMode: 'http-cache',
  phase: 'idle',
  currentChunk: null,
  currentChunkIndex: 0,
  currentChunkCount: 0,
  downloadedBytes: 0,
  downloadBytes: 329 * 1024 * 1024,
  verifiedBytes: 0,
  verifyBytes: 0,
  bytesPerSecond: 0,
  failedResources: [],
  canPause: false,
  canResume: false,
  canRetry: false,
};

afterEach(() => {
  document.body.replaceChildren();
});

describe('OfficeResourcePanel', () => {
  it.each(['zh-CN', 'en-US'] as const)(
    'renders the compact %s resource surface without responder warnings',
    async (locale) => {
      const listeners = new Set<() => void>();
      const manager = {
        getSnapshot: () => snapshot,
        subscribe: (listener: () => void) => {
          listeners.add(listener);
          return () => listeners.delete(listener);
        },
        prefetchRecommended: vi.fn(),
        installFontPreset: vi.fn(),
        repair: vi.fn(),
        loadAll: vi.fn(),
        pause: vi.fn(),
        resume: vi.fn(),
      } as unknown as OfficeRuntimeResourceManager;
      const error = vi.spyOn(console, 'error').mockImplementation(() => undefined);
      const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
      const host = document.createElement('div');
      document.body.append(host);
      const root = createRoot(host);

      await act(async () => {
        root.render(<OfficeResourcePanel manager={manager} copy={officeCopy[locale]} />);
      });

      expect(host.textContent).toContain('OnlyOffice 0.5.0');
      expect(host.textContent).toContain(officeCopy[locale].completePackage);
      expect(host.textContent).toContain(officeCopy[locale].packageIncludes);
      expect(host.textContent).not.toContain(officeCopy[locale].advancedFonts);
      expect(host.querySelectorAll('button').length).toBeGreaterThan(0);
      expect([...host.querySelectorAll('button')].every((button) => button.classList.contains('button--sm'))).toBe(
        true,
      );
      expect(host.querySelector('[role="progressbar"]')).toBeNull();
      expect(error).not.toHaveBeenCalled();
      expect(warn).not.toHaveBeenCalled();

      await act(async () => root.unmount());
      error.mockRestore();
      warn.mockRestore();
    },
  );

  it('shows the upstream package version when it is newer than the cached shell', async () => {
    const upstreamSnapshot = {
      ...snapshot,
      packageVersion: '0.5.5',
      availablePackageVersion: '0.5.6',
    };
    const manager = {
      getSnapshot: () => upstreamSnapshot,
      subscribe: () => () => undefined,
      repair: vi.fn(),
    } as unknown as OfficeRuntimeResourceManager;
    const host = document.createElement('div');
    document.body.append(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<OfficeResourcePanel manager={manager} copy={officeCopy['zh-CN']} />);
    });

    expect(host.textContent).toContain('OnlyOffice 0.5.6');
    expect(host.textContent).not.toContain('OnlyOffice 0.5.5');
    await act(async () => root.unmount());
  });

  it.each(['zh-CN', 'en-US'] as const)('shows the explicit %s download stage and package segment', async (locale) => {
    const downloadingSnapshot: OfficeRuntimeResourceSnapshot = {
      ...snapshot,
      readiness: 'updating',
      phase: 'downloading',
      operation: 'load-all',
      currentChunk: 'segment-006',
      currentChunkIndex: 6,
      currentChunkCount: 24,
      downloadedBytes: 141 * 1024 * 1024,
      downloadBytes: 594 * 1024 * 1024,
      verifiedBytes: 120 * 1024 * 1024,
      verifyBytes: 594 * 1024 * 1024,
      canPause: true,
    };
    const manager = {
      getSnapshot: () => downloadingSnapshot,
      subscribe: () => () => undefined,
      pause: vi.fn(),
    } as unknown as OfficeRuntimeResourceManager;
    const host = document.createElement('div');
    document.body.append(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<OfficeResourcePanel manager={manager} copy={officeCopy[locale]} />);
    });

    expect(host.textContent).toContain(`${officeCopy[locale].resourceStage} 2/4`);
    expect(host.textContent).toContain(officeCopy[locale].downloading);
    expect(host.textContent).toContain(`${officeCopy[locale].packageSegment} 6/24`);
    expect(host.textContent).toContain('141 MB / 594 MB');
    const progress = host.querySelector('[role="progressbar"]');
    expect(progress?.getAttribute('aria-valuetext')).toContain(officeCopy[locale].downloading);
    expect(progress?.getAttribute('aria-valuenow')).toBe(String(141 * 1024 * 1024));

    await act(async () => root.unmount());
  });

  it('uses verified bytes instead of downloaded bytes during the verification stage', async () => {
    const verifyingSnapshot: OfficeRuntimeResourceSnapshot = {
      ...snapshot,
      readiness: 'updating',
      phase: 'verifying',
      operation: 'load-all',
      currentChunk: 'segment-013',
      currentChunkIndex: 13,
      currentChunkCount: 24,
      downloadedBytes: 594 * 1024 * 1024,
      downloadBytes: 594 * 1024 * 1024,
      verifiedBytes: 300 * 1024 * 1024,
      verifyBytes: 594 * 1024 * 1024,
    };
    const manager = {
      getSnapshot: () => verifyingSnapshot,
      subscribe: () => () => undefined,
    } as unknown as OfficeRuntimeResourceManager;
    const host = document.createElement('div');
    document.body.append(host);
    const root = createRoot(host);

    await act(async () => {
      root.render(<OfficeResourcePanel manager={manager} copy={officeCopy['zh-CN']} />);
    });

    expect(host.textContent).toContain('阶段 3/4');
    expect(host.textContent).toContain('正在校验');
    expect(host.textContent).toContain('300 MB / 594 MB');
    expect(host.textContent).not.toContain('594 MB / 594 MB');

    await act(async () => root.unmount());
  });

  it.each(['zh-CN', 'en-US'] as const)(
    'offers a targeted %s repair when durable resources exist but the editor HTTP cache is cold',
    async (locale) => {
      const repair = vi.fn(async () => undefined);
      const repairSnapshot: OfficeRuntimeResourceSnapshot = {
        ...snapshot,
        readiness: 'repair-needed',
        error: { code: 'storage', path: 'office-resources.oobpack?segment=cold' },
      };
      const manager = {
        getSnapshot: () => repairSnapshot,
        subscribe: () => () => undefined,
        repair,
      } as unknown as OfficeRuntimeResourceManager;
      const host = document.createElement('div');
      document.body.append(host);
      const root = createRoot(host);

      await act(async () => {
        root.render(<OfficeResourcePanel manager={manager} copy={officeCopy[locale]} />);
      });

      const button = [...host.querySelectorAll('button')].find(
        (candidate) => candidate.textContent === officeCopy[locale].repair,
      );
      expect(button).toBeTruthy();
      await act(async () => {
        button?.click();
      });
      expect(repair).toHaveBeenCalledWith({ scope: 'all' });

      await act(async () => root.unmount());
    },
  );
});
