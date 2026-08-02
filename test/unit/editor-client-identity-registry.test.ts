import { describe, expect, it } from 'vitest';
import { EditorClientIdentityRegistry } from '../../src/lib/editor-client-identity-registry';

const sessionA = { releaseId: 'release-a', sessionId: 'session-a' };
const sessionB = { releaseId: 'release-b', sessionId: 'session-b' };

describe('EditorClientIdentityRegistry', () => {
  it('propagates the exact host identity through nested navigation client ids', () => {
    const registry = new EditorClientIdentityRegistry();
    registry.bindHost('host', sessionA);

    expect(
      registry.resolve({
        clientId: 'host',
        resultingClientId: 'frame-editor',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: false,
      }),
    ).toEqual(sessionA);
    expect(
      registry.resolve({
        clientId: 'frame-editor',
        resultingClientId: 'document-frame',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: true,
      }),
    ).toEqual(sessionA);
    expect(
      registry.resolve({
        clientId: 'document-frame',
        releaseId: null,
        sourceIsOfficeRuntime: true,
      }),
    ).toEqual(sessionA);
  });

  it('recovers a runtime child after Service Worker termination only from one exact host identity', () => {
    const registry = new EditorClientIdentityRegistry();
    expect(
      registry.resolve({
        clientId: 'runtime-child',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: true,
        recoveryIdentities: [sessionA],
      }),
    ).toEqual(sessionA);
    expect(
      new EditorClientIdentityRegistry().resolve({
        clientId: 'runtime-child',
        releaseId: null,
        sourceIsOfficeRuntime: true,
        recoveryIdentities: [sessionA, sessionB],
      }),
    ).toBeNull();
  });

  it('does not let an unrelated same-origin window inherit a connected capability', () => {
    const registry = new EditorClientIdentityRegistry();
    expect(
      registry.resolve({
        clientId: 'settings-page',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: false,
        connectedIdentity: sessionA,
        recoveryIdentities: [sessionA],
      }),
    ).toBeNull();
  });

  it('does not let an unknown runtime child inherit A while live A and B Host identities coexist', () => {
    const registry = new EditorClientIdentityRegistry();
    registry.bindHost('host-a', sessionA);
    expect(
      registry.resolve({
        clientId: 'runtime-child-b',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: true,
        connectedIdentity: sessionA,
        recoveryIdentities: [sessionA, sessionB],
      }),
    ).toBeNull();
    expect(
      registry.resolve({
        clientId: 'runtime-child-a',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: true,
        connectedIdentity: sessionA,
        recoveryIdentities: [sessionA],
      }),
    ).toEqual(sessionA);
  });

  it('clears retired client lineage when an origin slot binds a new session', () => {
    const registry = new EditorClientIdentityRegistry();
    registry.bindHost('host-a', sessionA);
    registry.resolve({
      clientId: 'host-a',
      resultingClientId: 'old-child',
      releaseId: 'release-a',
      sourceIsOfficeRuntime: false,
    });
    registry.bindHost('host-b', sessionB);

    expect(
      registry.resolve({
        clientId: 'old-child',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: false,
      }),
    ).toBeNull();
    expect(
      registry.resolve({
        clientId: 'host-b',
        releaseId: 'release-b',
        sourceIsOfficeRuntime: false,
      }),
    ).toEqual(sessionB);
  });

  it('removes the exact host and nested lineage on an authenticated unbind', () => {
    const registry = new EditorClientIdentityRegistry();
    registry.bindHost('host-a', sessionA);
    registry.resolve({
      clientId: 'host-a',
      resultingClientId: 'runtime-child',
      releaseId: 'release-a',
      sourceIsOfficeRuntime: false,
    });

    expect(registry.unbindHost('host-a', sessionB)).toBe(false);
    expect(registry.unbindHost('unknown', sessionA)).toBe(false);
    expect(registry.unbindHost('host-a', sessionA)).toBe(true);
    expect(
      registry.resolve({
        clientId: 'runtime-child',
        releaseId: 'release-a',
        sourceIsOfficeRuntime: false,
      }),
    ).toBeNull();
  });
});
