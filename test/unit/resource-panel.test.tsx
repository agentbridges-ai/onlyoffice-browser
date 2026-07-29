import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { OfficeRuntimeResourceManager, OfficeRuntimeResourceSnapshot } from '../../src/lib/runtime-resources';
import { officeCopy } from '../../src/pwa/i18n';
import { OfficeResourcePanel } from '../../src/pwa/resource-panel';

(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const snapshot: OfficeRuntimeResourceSnapshot = {
  packageVersion: '0.4.2',
  assetVersion: 'v0.4.2-test',
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
  installedRelease: 'v0.4.2-test',
  targetRelease: 'v0.4.2-test',
  availableRelease: 'v0.4.2-test',
  storageMode: 'http-cache',
  phase: 'idle',
  currentChunk: null,
  downloadedBytes: 0,
  downloadBytes: 0,
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

      expect(host.textContent).toContain('OnlyOffice 0.4.2');
      expect(host.textContent).toContain(officeCopy[locale].advancedFonts);
      expect(host.querySelectorAll('button').length).toBeGreaterThan(0);
      expect(error).not.toHaveBeenCalled();
      expect(warn).not.toHaveBeenCalled();

      await act(async () => root.unmount());
      error.mockRestore();
      warn.mockRestore();
    },
  );
});
