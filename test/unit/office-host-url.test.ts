import { describe, expect, it } from 'vitest';
import {
  readOfficeHostBootstrap,
  resolveOriginBoundWorkerUrl,
  writeOfficeHostBootstrap,
} from '../../src/lib/office-host-url';

describe('Office Host URL', () => {
  it('keeps per-session bootstrap values out of the HTTP cache key', () => {
    const first = writeOfficeHostBootstrap(new URL('https://aries.getpi.work/r/release/office-host.html'), {
      sessionId: 'office-editor-first',
      parentOrigin: 'https://onlyoffice.getpi.work',
    });
    const second = writeOfficeHostBootstrap(new URL('https://aries.getpi.work/r/release/office-host.html'), {
      sessionId: 'office-editor-second',
      parentOrigin: 'https://onlyoffice.getpi.work',
    });

    expect(`${first.origin}${first.pathname}${first.search}`).toBe(
      `${second.origin}${second.pathname}${second.search}`,
    );
    expect(first.search).toBe('');
    expect(second.search).toBe('');
    expect(readOfficeHostBootstrap(first)).toEqual({
      sessionId: 'office-editor-first',
      parentOrigin: 'https://onlyoffice.getpi.work',
      releaseId: 'release',
    });
    expect(readOfficeHostBootstrap(second)).toEqual({
      sessionId: 'office-editor-second',
      parentOrigin: 'https://onlyoffice.getpi.work',
      releaseId: 'release',
    });
  });

  it('continues to read query bootstrap values from older npm clients', () => {
    expect(
      readOfficeHostBootstrap(
        new URL('https://aries.getpi.work/office-host.html?sessionId=legacy&parentOrigin=https%3A%2F%2Fpiwork.example'),
      ),
    ).toEqual({
      sessionId: 'legacy',
      parentOrigin: 'https://piwork.example',
      releaseId: '',
    });
  });

  it('binds generated worker paths to the current isolated editor origin', () => {
    const location = {
      origin: 'https://aries.getpi.work',
    } as Location;
    expect(
      resolveOriginBoundWorkerUrl(
        new URL('https://onlyoffice.getpi.work/r/release/wasm/x2t/startup-heartbeat-worker-abcd.js'),
        location,
      ).href,
    ).toBe('https://aries.getpi.work/r/release/wasm/x2t/startup-heartbeat-worker-abcd.js');
  });
});
