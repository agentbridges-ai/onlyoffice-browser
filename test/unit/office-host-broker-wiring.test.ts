import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

const source = readFileSync(resolve(process.cwd(), 'src/office-host.ts'), 'utf8');

describe('Office Host canonical Resource Broker wiring', () => {
  it('pins the bootstrap release, session and parent before announcing readiness', () => {
    expect(source).toContain(
      'const { sessionId, parentOrigin, releaseId } = readOfficeHostBootstrap(new URL(window.location.href));',
    );

    const announceStart = source.indexOf('async function announceHostReady()');
    const ensureWorker = source.indexOf('await ensureOfficeServiceWorkerControl();', announceStart);
    const bindBroker = source.indexOf('await ensureCanonicalResourceBrokerBound();', announceStart);
    const loadIdentity = source.indexOf('const identity = await loadOfficeHostIdentity();', announceStart);
    const postReady = source.indexOf("type: 'HOST_READY'", announceStart);

    expect(announceStart).toBeGreaterThan(-1);
    expect(ensureWorker).toBeGreaterThan(announceStart);
    expect(bindBroker).toBeGreaterThan(ensureWorker);
    expect(loadIdentity).toBeGreaterThan(bindBroker);
    expect(postReady).toBeGreaterThan(loadIdentity);
  });

  it('transfers exactly the canonical broker and reply ports using the strict pinned protocol', () => {
    expect(source).toContain('type: EDITOR_RESOURCE_BROKER_BIND_TYPE');
    expect(source).toContain('[connection.port, reply.port2]');
    expect(source).toContain("hasExactKeys(value, ['protocol', 'type', 'ok', 'releaseId', 'sessionId'])");
    expect(source).toContain('value.releaseId === releaseId');
    expect(source).toContain('value.sessionId === sessionId');
    expect(source).toContain('navigator.serviceWorker.controller !== controller');
    expect(source).not.toContain("postMessage('*'");
    expect(source).not.toContain('postMessage("*"');
  });

  it('fails closed for versioned hosts and restores the broker after worker or port loss', () => {
    expect(source).toContain("throw new Error('Service workers are required for a versioned Office editor host')");
    expect(source).toContain(
      "return Promise.reject(new Error('Office service worker is not controlling the versioned editor host'))",
    );
    expect(source).toContain("type: 'ONLYOFFICE_BROKER_NEEDED'");
    expect(source).toContain('event.source !== navigator.serviceWorker.controller');
    expect(source).toContain("navigator.serviceWorker.addEventListener('message', handleResourceBrokerNeeded)");
    expect(source).toContain(
      "navigator.serviceWorker.addEventListener('controllerchange', handleOfficeServiceWorkerControllerChange)",
    );
    expect(source).toContain('void ensureCanonicalResourceBrokerBound(true)');
  });

  it('keeps local unversioned development compatible and cleans up the hidden broker frame', () => {
    expect(source).toContain('return Boolean(releaseId);');
    expect(source).toContain("canonical.hostname = 'onlyoffice.localhost'");
    expect(source).toContain('allowLocalMatrix: true');
    expect(source).toContain('resourceBrokerFrameClient?.destroy()');
    expect(source).toContain('disposeResourceBroker();');
    expect(source).toContain('type: EDITOR_RESOURCE_BROKER_UNBIND_TYPE');
    expect(source).toContain('navigator.serviceWorker.controller?.postMessage({');
    expect(source).toContain(
      "navigator.serviceWorker.removeEventListener('controllerchange', handleOfficeServiceWorkerControllerChange)",
    );
  });
});
