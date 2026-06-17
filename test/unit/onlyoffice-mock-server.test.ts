import { describe, expect, it, vi } from 'vitest';

import {
  createOnlyOfficeMockServer,
  getOnlyOfficeMediaCandidates,
  resolveOnlyOfficeMediaUrl,
} from '../../src/lib/onlyoffice-mock-server';

describe('onlyoffice-mock-server', () => {
  it('returns a single local participant for browser-only editing', () => {
    const server = createOnlyOfficeMockServer();

    expect(server.getInitialChanges?.()).toEqual([]);
    expect(server.getParticipants()).toEqual({
      index: 0,
      list: [
        {
          id: 1,
          idOriginal: 'local-browser-user',
          username: 'Guest',
          indexUser: 0,
          connectionId: 'local-browser-session',
          isCloseCoAuthoring: false,
          view: false,
        },
      ],
    });
  });

  it('resolves media names without leaving the local object URL map', async () => {
    const media = {
      'media/image 1.png': 'blob:image-1',
      'media/chart.png': 'blob:chart',
    };
    const server = createOnlyOfficeMockServer({ media });

    expect(getOnlyOfficeMediaCandidates('./image%201.png')).toContain('media/image 1.png');
    expect(resolveOnlyOfficeMediaUrl(media, 'chart.png')).toBe('blob:chart');
    await expect(server.getImageURL?.('./image%201.png')).resolves.toBe('blob:image-1');
    await expect(server.getImageURL?.('missing.png')).resolves.toBe('');
  });

  it('passes OnlyOffice messages to the optional observer', () => {
    const onMessage = vi.fn();
    const server = createOnlyOfficeMockServer({ onMessage });
    const msg = { type: 'auth', openCmd: { url: 'blob:doc' } };

    server.onMessage(msg);

    expect(onMessage).toHaveBeenCalledWith(msg);
  });

  it('exposes build metadata for the CryptPad wrapper auth response', () => {
    const server = createOnlyOfficeMockServer({ buildVersion: '9.3.0', buildNumber: 140 });

    expect(server.buildVersion).toBe('9.3.0');
    expect(server.buildNumber).toBe(140);
  });
});
