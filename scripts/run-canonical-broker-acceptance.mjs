#!/usr/bin/env node

import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  CdpClient,
  attachToTarget,
  evaluate,
  findOrCreateTarget,
  waitForExpression,
} from './chrome-cdp-memory-stress.mjs';
import {
  analyzeBrokerMetrics,
  analyzeLifecycleReport,
  analyzeRecoveryResults,
  analyzeReleaseIntegrity,
  analyzeSecurityProbes,
  analyzeStartupPerformance,
  analyzeTrafficEvidence,
  combineAcceptanceSections,
  percentile,
} from './canonical-broker-acceptance-policy.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_BROWSER_WS = 'ws://127.0.0.1:9222/devtools/browser';
const DEFAULT_APP_URL = 'https://onlyoffice.getpi.work/';
const DEFAULT_TIMEOUT_MS = 180_000;
const REQUIRED_ITERATIONS = 100;
const REQUIRED_ORIGINS = 3;
const RECOVERY_TIMEOUT_MS = 30_000;
const BROKER_PROTOCOL = 'onlyoffice-browser-resource-broker/v1';
const BROKER_WINDOW_BYTES = 256 * 1024;
const SHA256_PATTERN = /^[a-f0-9]{64}$/u;

function parseArgs(argv) {
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const options = {
    browserWs: DEFAULT_BROWSER_WS,
    url: DEFAULT_APP_URL,
    hostUrl: null,
    profileDir: null,
    baselineP95Ms: null,
    baselineReport: null,
    timeoutMs: DEFAULT_TIMEOUT_MS,
    sections: new Set(['lifecycle', 'concurrency', 'recovery', 'integrity', 'security']),
    trafficEvidenceUrl: null,
    output: resolve(ROOT, 'test-results', 'broker-acceptance', `canonical-broker-${timestamp}.json`),
    lifecycleJson: null,
    lifecycleCsv: null,
    keepTargets: false,
  };
  for (const arg of argv) {
    if (arg.startsWith('--browser-ws=')) options.browserWs = arg.slice('--browser-ws='.length);
    else if (arg.startsWith('--url=')) options.url = arg.slice('--url='.length);
    else if (arg.startsWith('--host-url=')) options.hostUrl = arg.slice('--host-url='.length);
    else if (arg.startsWith('--profile-dir=')) options.profileDir = resolve(arg.slice('--profile-dir='.length));
    else if (arg.startsWith('--baseline-p95-ms=')) {
      options.baselineP95Ms = Number(arg.slice('--baseline-p95-ms='.length));
    } else if (arg.startsWith('--baseline-report=')) {
      options.baselineReport = resolve(arg.slice('--baseline-report='.length));
    } else if (arg.startsWith('--timeout-ms=')) {
      options.timeoutMs = Number(arg.slice('--timeout-ms='.length));
    } else if (arg.startsWith('--sections=')) {
      options.sections = new Set(
        arg
          .slice('--sections='.length)
          .split(',')
          .map((value) => value.trim())
          .filter(Boolean),
      );
    } else if (arg.startsWith('--traffic-evidence-url=')) {
      options.trafficEvidenceUrl = arg.slice('--traffic-evidence-url='.length);
    } else if (arg.startsWith('--output=')) options.output = resolve(arg.slice('--output='.length));
    else if (arg.startsWith('--lifecycle-json=')) {
      options.lifecycleJson = resolve(arg.slice('--lifecycle-json='.length));
    } else if (arg.startsWith('--lifecycle-csv=')) {
      options.lifecycleCsv = resolve(arg.slice('--lifecycle-csv='.length));
    } else if (arg === '--keep-targets') options.keepTargets = true;
    else if (arg === '--help' || arg === '-h') {
      printHelp();
      process.exit(0);
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }
  const allowedSections = new Set(['lifecycle', 'concurrency', 'recovery', 'integrity', 'security']);
  if (!options.sections.size || [...options.sections].some((section) => !allowedSections.has(section))) {
    throw new Error('--sections must contain one or more of: lifecycle,concurrency,recovery,integrity,security');
  }
  if (!options.profileDir) {
    throw new Error(
      '--profile-dir is required; acceptance must reuse and verify an already-installed persistent profile',
    );
  }
  if (!existsSync(options.profileDir)) throw new Error(`Persistent profile does not exist: ${options.profileDir}`);
  if (options.sections.has('lifecycle') && !Number.isFinite(options.baselineP95Ms) && !options.baselineReport) {
    throw new Error('--baseline-p95-ms or --baseline-report is required before the lifecycle run starts');
  }
  if (!Number.isFinite(options.timeoutMs) || options.timeoutMs < RECOVERY_TIMEOUT_MS) {
    throw new Error(`--timeout-ms must be at least ${RECOVERY_TIMEOUT_MS}`);
  }
  const outputDir = dirname(options.output);
  options.trafficEvidenceUrl ||= new URL('/__matrix__/content-counters', options.url).href;
  options.lifecycleJson ??= resolve(outputDir, 'lifecycle.json');
  options.lifecycleCsv ??= resolve(outputDir, 'lifecycle.csv');
  return options;
}

function printHelp() {
  console.log(`Canonical Cache Broker production acceptance

This command never installs Office resources. It attaches to an existing Chrome
whose persistent profile already contains a verified canonical installation.

Required:
  --profile-dir=/path/to/chrome-profile
  --baseline-p95-ms=123
    or --baseline-report=/path/to/direct-http-cache-baseline.json

Optional:
  --browser-ws=ws://127.0.0.1:9222/devtools/browser
  --url=https://onlyoffice.getpi.work/
  --host-url=https://{hostSlot}.getpi.work/office-host.html
  --sections=lifecycle,concurrency,recovery,integrity,security
  --traffic-evidence-url=https://metrics.example/content-counters
  --output=/path/to/auditable-report.json
  --keep-targets

The fixed production gates are printed and recorded before work starts:
  lifecycle=100 cycles, concurrency=3 origins, recovery<=30s,
  Broker high-water<=64MiB, startup extra P95<=max(300ms, baseline*10%),
  every manifest asset exact through Broker, and zero Worker/R2 counter delta.
`);
}

function sha256File(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

function resolveBaseline(options) {
  if (Number.isFinite(options.baselineP95Ms)) {
    return { p95Ms: options.baselineP95Ms, source: 'argument', reportPath: null, reportSha256: null };
  }
  if (!options.baselineReport || !existsSync(options.baselineReport)) {
    throw new Error(`Baseline report does not exist: ${options.baselineReport}`);
  }
  const report = JSON.parse(readFileSync(options.baselineReport, 'utf8'));
  const readyP95Ms = percentile(
    (report.samples || []).filter((sample) => sample.phase === 'closed').map((sample) => sample.readyWaitMs),
    0.95,
  );
  if (!Number.isFinite(readyP95Ms)) throw new Error('Baseline report has no finite closed-state ready samples');
  return {
    p95Ms: readyP95Ms,
    source: 'report',
    reportPath: options.baselineReport,
    reportSha256: sha256File(options.baselineReport),
  };
}

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

function releaseAssetUrl(origin, releaseId, path) {
  return new URL(
    `/r/${encodeURIComponent(releaseId)}/${path
      .split('/')
      .map((part) => encodeURIComponent(part))
      .join('/')}`,
    origin,
  ).href;
}

function validManifestAssetPath(value) {
  if (
    typeof value !== 'string' ||
    !value ||
    value.length > 2_048 ||
    value.startsWith('/') ||
    value.includes('\\') ||
    value.includes('?') ||
    value.includes('#') ||
    /^[a-zA-Z][a-zA-Z0-9+.-]*:/u.test(value) ||
    [...value].some((character) => {
      const code = character.charCodeAt(0);
      return code <= 0x1f || code === 0x7f;
    })
  ) {
    return false;
  }
  try {
    if (decodeURIComponent(value) !== value || value.normalize('NFC') !== value) return false;
  } catch {
    return false;
  }
  return value.split('/').every((segment) => segment && segment !== '.' && segment !== '..');
}

function isEditorShellAssetPath(path) {
  return (
    path === 'office-host.html' ||
    path === 'editor-shell-prime.html' ||
    /^assets\/[a-zA-Z0-9._+-]+\.(?:css|js)$/u.test(path)
  );
}

async function loadReleaseManifest(options, releaseId) {
  const url = new URL(`/releases/${encodeURIComponent(releaseId)}/manifest.json`, options.url).href;
  const response = await fetch(url, { cache: 'no-store', headers: { Accept: 'application/json' } });
  if (!response.ok) throw new Error(`Release manifest request failed: ${response.status} ${url}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const manifestSha256 = createHash('sha256').update(bytes).digest('hex');
  const manifest = JSON.parse(new TextDecoder().decode(bytes));
  if (manifest?.version !== 5 || manifest.releaseId !== releaseId || !Array.isArray(manifest.assets)) {
    throw new Error('Release manifest must be v5, match the installed release, and contain assets');
  }
  if (!manifest.assets.length) throw new Error('Release manifest has no assets');
  const paths = new Set();
  const assets = manifest.assets.map((asset) => {
    if (
      !asset ||
      !validManifestAssetPath(asset.path) ||
      paths.has(asset.path) ||
      !Number.isSafeInteger(asset.bytes) ||
      asset.bytes < 0 ||
      typeof asset.mime !== 'string' ||
      !asset.mime ||
      !SHA256_PATTERN.test(asset.sha256)
    ) {
      throw new Error(`Release manifest contains an invalid asset: ${asset?.path ?? '<unknown>'}`);
    }
    paths.add(asset.path);
    const fastCdcBoundaries = Array.isArray(asset.representations?.fastcdc?.chunks)
      ? asset.representations.fastcdc.chunks
          .map((chunk) => chunk?.offset)
          .filter((offset) => Number.isSafeInteger(offset) && offset > 0 && offset < asset.bytes)
          .sort((left, right) => left - right)
      : [];
    return {
      path: asset.path,
      bytes: asset.bytes,
      mime: asset.mime,
      sha256: asset.sha256,
      fastCdcBoundaries,
    };
  });
  return { url, manifestSha256, releaseId, assets };
}

async function fetchTrafficEvidence(options) {
  const headers = { Accept: 'application/json' };
  const bearer = process.env.ONLYOFFICE_TRAFFIC_EVIDENCE_BEARER;
  if (bearer) headers.Authorization = `Bearer ${bearer}`;
  const response = await fetch(options.trafficEvidenceUrl, { cache: 'no-store', headers });
  if (!response.ok) {
    throw new Error(
      `HTTP/R2 traffic evidence interface is required and returned ${response.status}: ${options.trafficEvidenceUrl}`,
    );
  }
  const value = await response.json().catch(() => null);
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`HTTP/R2 traffic evidence interface returned invalid JSON: ${options.trafficEvidenceUrl}`);
  }
  return value;
}

async function runCommand(command, args) {
  await new Promise((resolvePromise, reject) => {
    const child = spawn(command, args, {
      cwd: ROOT,
      env: process.env,
      stdio: 'inherit',
    });
    child.once('error', reject);
    child.once('exit', (code, signal) => {
      if (code === 0) resolvePromise();
      else reject(new Error(`${command} exited with ${code ?? signal}`));
    });
  });
}

async function preflight(cdp, options) {
  const browserVersion = await cdp.send('Browser.getVersion');
  const commandLine = await cdp.send('Browser.getBrowserCommandLine').catch(() => null);
  const expectedProfile = resolve(options.profileDir);
  const profileArgumentIndex =
    commandLine?.arguments?.findIndex(
      (argument) => argument === '--user-data-dir' || argument.startsWith('--user-data-dir='),
    ) ?? -1;
  const profileArgument = profileArgumentIndex >= 0 ? commandLine.arguments[profileArgumentIndex] : null;
  const profileValue =
    profileArgument === '--user-data-dir'
      ? commandLine.arguments[profileArgumentIndex + 1]
      : profileArgument?.slice('--user-data-dir='.length);
  const actualProfile = profileValue ? resolve(profileValue) : null;
  if (actualProfile !== expectedProfile) {
    throw new Error(
      `Connected Chrome profile mismatch: expected ${expectedProfile}, got ${actualProfile || 'unavailable'}. ` +
        'Start Chrome with --enable-automation and the exact --user-data-dir before acceptance.',
    );
  }

  const target = await findOrCreateTarget(cdp, options.url, true);
  const sessionId = await attachToTarget(cdp, target.targetId);
  await cdp.send('Page.navigate', { url: options.url }, sessionId);
  await waitForExpression(
    cdp,
    sessionId,
    `Boolean(window.__officeDemo?.resourceSnapshot)`,
    options.timeoutMs,
    'the resource manager snapshot',
  );
  const snapshot = await evaluate(
    cdp,
    sessionId,
    `(() => {
      const value = window.__officeDemo?.resourceSnapshot;
      return value ? JSON.parse(JSON.stringify(value)) : null;
    })()`,
  );
  if (snapshot?.readiness !== 'ready') {
    throw new Error(
      `Office resources are not ready (${snapshot?.readiness || 'missing'}); acceptance will not install them`,
    );
  }
  const documentState = await evaluate(
    cdp,
    sessionId,
    `(() => ({
      tabCount: window.__officeDemo?.tabs?.length ?? 0,
      activeEditor: Boolean(window.__officeDemo?.editor),
      dirtyCount: Array.from(window.__officeDemo?.tabs ?? []).filter((tab) => tab.dirty).length,
    }))()`,
  );
  if (documentState.tabCount !== 0 || documentState.activeEditor) {
    throw new Error(
      `Acceptance profile has ${documentState.tabCount} open document tab(s) ` +
        `(${documentState.dirtyCount} dirty); close them manually before the run`,
    );
  }
  const releaseId = snapshot.installedRelease || snapshot.targetRelease;
  if (typeof releaseId !== 'string' || !releaseId) throw new Error('Ready snapshot has no installed release identity');
  await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
  return {
    browserVersion,
    profileDir: actualProfile,
    persistentProfileVerified: true,
    documentState,
    resourceSnapshot: snapshot,
    releaseId,
    targetId: target.targetId,
    targetCreatedByRunner: target.createdByScript,
  };
}

async function clearBrowserHttpCache(cdp, targetId) {
  const sessionId = await attachToTarget(cdp, targetId);
  try {
    await cdp.send('Network.enable', {}, sessionId);
    await cdp.send('Network.clearBrowserCache', {}, sessionId);
    return {
      cleared: true,
      method: 'Network.clearBrowserCache',
      completedAt: new Date().toISOString(),
      preservedStores: ['Cache Storage', 'IndexedDB'],
    };
  } finally {
    await cdp.send('Network.disable', {}, sessionId).catch(() => {});
    await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
  }
}

async function runLifecycle(options, baseline) {
  const args = [
    resolve(ROOT, 'scripts', 'chrome-cdp-memory-stress.mjs'),
    `--browser-ws=${options.browserWs}`,
    `--url=${options.url}`,
    `--iterations=${REQUIRED_ITERATIONS}`,
    '--formats=docx,xlsx,pptx',
    '--modes=edit',
    '--batch-size=1',
    '--stay-ms=0',
    '--home-dwell-ms=0',
    '--opened-sample-interval=10',
    '--final-settle-ms=30000',
    '--close-mode=direct',
    '--keep-target',
    `--json-output=${options.lifecycleJson}`,
    `--csv-output=${options.lifecycleCsv}`,
  ];
  if (options.hostUrl) args.push(`--host-url=${options.hostUrl}`);
  await runCommand(process.execPath, args);
  const report = JSON.parse(readFileSync(options.lifecycleJson, 'utf8'));
  return {
    reportPath: options.lifecycleJson,
    reportSha256: sha256File(options.lifecycleJson),
    csvPath: options.lifecycleCsv,
    csvSha256: sha256File(options.lifecycleCsv),
    policy: analyzeLifecycleReport(report),
    startup: analyzeStartupPerformance(report, baseline.p95Ms),
  };
}

async function evaluateTarget(cdp, target, expression) {
  const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: target.targetId, flatten: true });
  try {
    await cdp.send('Runtime.enable', {}, sessionId);
    return await evaluate(cdp, sessionId, expression);
  } finally {
    await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
  }
}

async function evaluateInContext(cdp, context, expression) {
  const result = await cdp.send(
    'Runtime.evaluate',
    {
      expression,
      contextId: context.contextId,
      awaitPromise: true,
      returnByValue: true,
    },
    context.sessionId,
  );
  if (result.exceptionDetails) {
    const description =
      result.exceptionDetails.exception?.description ||
      result.exceptionDetails.text ||
      'Unknown Runtime.evaluate exception';
    throw new Error(description);
  }
  return result.result?.value;
}

async function findExecutionContext(cdp, predicate, label, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  let lastUrls = [];
  while (Date.now() < deadline) {
    const { targetInfos } = await cdp.send('Target.getTargets');
    const targets = targetInfos.filter((target) => ['page', 'iframe'].includes(target.type));
    lastUrls = [];
    for (const target of targets) {
      const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: target.targetId, flatten: true });
      const contexts = [];
      cdp.on('Runtime.executionContextCreated', (event, eventSessionId) => {
        if (eventSessionId === sessionId && event.context?.auxData?.isDefault) contexts.push(event.context);
      });
      try {
        await cdp.send('Runtime.disable', {}, sessionId).catch(() => {});
        await cdp.send('Runtime.enable', {}, sessionId);
        await sleep(50);
        for (const context of contexts) {
          const locationResult = await cdp
            .send(
              'Runtime.evaluate',
              {
                expression: 'location.href',
                contextId: context.id,
                returnByValue: true,
              },
              sessionId,
            )
            .catch(() => null);
          const url = locationResult?.result?.value;
          if (typeof url !== 'string') continue;
          lastUrls.push(url);
          if (predicate(url)) {
            return {
              sessionId,
              contextId: context.id,
              targetId: target.targetId,
              url,
            };
          }
        }
      } catch {
        // Cross-process frames can disappear while targets are enumerated.
      }
      await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
    }
    await sleep(100);
  }
  throw new Error(`Unable to find ${label} execution context; observed ${lastUrls.join(', ') || 'no frame URLs'}`);
}

async function detachExecutionContext(cdp, context) {
  if (context?.sessionId) {
    await cdp.send('Target.detachFromTarget', { sessionId: context.sessionId }).catch(() => {});
  }
}

let brokerReadSequence = 0;

async function streamBrokerRead(cdp, context, request, options = {}) {
  const startedAt = Date.now();
  brokerReadSequence += 1;
  const bindingName = `__onlyofficeBrokerChunk${brokerReadSequence}`;
  const ackName = `__onlyofficeBrokerAck${brokerReadSequence}`;
  const bindingId = `broker-read-${brokerReadSequence}`;
  const hash = createHash('sha256');
  const captures = (options.captureRanges || []).map((capture) => {
    if (
      typeof capture?.id !== 'string' ||
      !capture.id ||
      !Number.isSafeInteger(capture.start) ||
      !Number.isSafeInteger(capture.end) ||
      capture.start < 0 ||
      capture.end < capture.start
    ) {
      throw new TypeError('Invalid Broker capture range');
    }
    return { ...capture, hash: createHash('sha256'), bytes: 0 };
  });
  let boundBytes = 0;
  let bindingError = null;
  let terminalResolve;
  let terminalReject;
  const terminal = new Promise((resolvePromise, reject) => {
    terminalResolve = resolvePromise;
    terminalReject = reject;
  });
  cdp.on('Runtime.bindingCalled', (event, eventSessionId) => {
    if (
      eventSessionId !== context.sessionId ||
      event.executionContextId !== context.contextId ||
      event.name !== bindingName
    ) {
      return;
    }
    try {
      const payload = JSON.parse(event.payload);
      if (payload.id !== bindingId) return;
      if (payload.type === 'chunk') {
        const bytes = Buffer.from(payload.base64, 'base64');
        const chunkStart = boundBytes;
        const chunkEnd = chunkStart + bytes.byteLength - 1;
        for (const capture of captures) {
          const overlapStart = Math.max(chunkStart, capture.start);
          const overlapEnd = Math.min(chunkEnd, capture.end);
          if (overlapEnd < overlapStart) continue;
          const slice = bytes.subarray(overlapStart - chunkStart, overlapEnd - chunkStart + 1);
          capture.hash.update(slice);
          capture.bytes += slice.byteLength;
        }
        hash.update(bytes);
        boundBytes += bytes.byteLength;
        void cdp
          .send(
            'Runtime.evaluate',
            {
              expression: `globalThis[${JSON.stringify(ackName)}]?.(${JSON.stringify(payload.sequence)})`,
              contextId: context.contextId,
              returnByValue: true,
            },
            context.sessionId,
          )
          .catch((error) => {
            bindingError = error;
            terminalReject(error);
          });
      } else if (payload.type === 'terminal') {
        terminalResolve();
      }
    } catch (error) {
      bindingError = error;
      terminalReject(error);
    }
  });
  await cdp.send('Runtime.addBinding', { name: bindingName, executionContextId: context.contextId }, context.sessionId);
  const hardTimeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const noProgressTimeoutMs = Math.min(hardTimeoutMs, RECOVERY_TIMEOUT_MS);
  const bindingHardTimeoutMs = hardTimeoutMs + 1_000;
  const config = {
    protocol: BROKER_PROTOCOL,
    request,
    bindingName,
    ackName,
    bindingId,
    noProgressTimeoutMs,
    cancelAfterFirstChunk: options.cancelAfterFirstChunk === true,
  };
  const evaluation = evaluateInContext(
    cdp,
    context,
    `(async () => {
      const config = ${JSON.stringify(config)};
      const emit = (value) => globalThis[config.bindingName](JSON.stringify({ id: config.bindingId, ...value }));
      const encode = (bytes) => {
        let binary = '';
        const step = 16 * 1024;
        for (let offset = 0; offset < bytes.length; offset += step) {
          const end = Math.min(bytes.length, offset + step);
          let part = '';
          for (let index = offset; index < end; index += 1) part += String.fromCharCode(bytes[index]);
          binary += part;
        }
        return btoa(binary);
      };
      const registration = await navigator.serviceWorker.ready;
      const worker = navigator.serviceWorker.controller ?? registration.active;
      if (!worker) throw new Error('Canonical Service Worker is unavailable');
      const channel = new MessageChannel();
      return await new Promise((resolve) => {
        let headers = null;
        let bytes = 0;
        let chunks = 0;
        let settled = false;
        let timer = null;
        let messageQueue = Promise.resolve();
        const finish = (value) => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          delete globalThis[config.ackName];
          channel.port1.close();
          emit({ type: 'terminal' });
          resolve({ headers, bytes, chunks, ...value });
        };
        const armTimeout = () => {
          clearTimeout(timer);
          timer = setTimeout(
            () => finish({ outcome: 'timeout-no-response', terminalType: null, errorCode: null }),
            config.noProgressTimeoutMs,
          );
        };
        armTimeout();
        channel.port1.onmessage = (event) => {
          armTimeout();
          messageQueue = messageQueue
            .then(async () => {
              const message = event.data;
              if (!message || message.protocol !== config.protocol || message.id !== config.request.id) {
                finish({ outcome: 'ERROR', terminalType: null, errorCode: 'protocol' });
                return;
              }
              if (message.type === 'HEADERS') {
                headers = { status: message.status, ...message.headers };
                if (message.status !== 416 && message.headers.contentLength > 0) {
                  channel.port1.postMessage({ protocol: config.protocol, type: 'PULL', id: config.request.id });
                }
                return;
              }
              if (message.type === 'CHUNK' && message.bytes instanceof ArrayBuffer) {
                const chunk = new Uint8Array(message.bytes);
                chunks += 1;
                const sequence = chunks;
                const acknowledged = new Promise((resolve, reject) => {
                  const timeout = setTimeout(() => {
                    delete globalThis[config.ackName];
                    reject(new Error('Node chunk acknowledgement timed out'));
                  }, config.noProgressTimeoutMs);
                  globalThis[config.ackName] = (receivedSequence) => {
                    if (receivedSequence !== sequence) return;
                    clearTimeout(timeout);
                    delete globalThis[config.ackName];
                    resolve();
                  };
                });
                emit({ type: 'chunk', sequence, base64: encode(chunk) });
                try {
                  await acknowledged;
                } catch {
                  finish({ outcome: 'ERROR', terminalType: null, errorCode: 'timeout' });
                  return;
                }
                if (settled) return;
                armTimeout();
                bytes += chunk.byteLength;
                channel.port1.postMessage({
                  protocol: config.protocol,
                  type: config.cancelAfterFirstChunk ? 'CANCEL' : 'PULL',
                  id: config.request.id,
                });
                return;
              }
              if (message.type === 'END' || message.type === 'CANCELLED') {
                finish({
                  outcome: message.type,
                  terminalType: message.type,
                  bytesSent: message.bytesSent,
                  errorCode: null,
                });
                return;
              }
              if (message.type === 'ERROR') {
                finish({ outcome: 'ERROR', terminalType: null, errorCode: message.code });
                return;
              }
              finish({ outcome: 'ERROR', terminalType: null, errorCode: 'protocol' });
            })
            .catch(() => {
              finish({ outcome: 'ERROR', terminalType: null, errorCode: 'protocol' });
            });
        };
        channel.port1.onmessageerror = () =>
          finish({ outcome: 'ERROR', terminalType: null, errorCode: 'protocol' });
        channel.port1.start();
        worker.postMessage(config.request, [channel.port2]);
      });
    })()`,
  );
  const [result] = await Promise.all([
    evaluation,
    Promise.race([
      terminal,
      new Promise((_, reject) =>
        setTimeout(
          () => reject(new Error(`Broker binding hard timeout elapsed for ${request.path ?? request.id}`)),
          bindingHardTimeoutMs,
        ),
      ),
    ]),
  ]);
  if (bindingError) throw bindingError;
  if (result.bytes !== boundBytes) {
    throw new Error(`Broker binding byte mismatch for ${request.path ?? request.id}: ${boundBytes} != ${result.bytes}`);
  }
  const captureResults = {};
  for (const capture of captures) {
    const expectedBytes = capture.end - capture.start + 1;
    if (capture.bytes !== expectedBytes) {
      throw new Error(`Broker capture ${capture.id} has ${capture.bytes} bytes; expected ${expectedBytes}`);
    }
    captureResults[capture.id] = {
      start: capture.start,
      end: capture.end,
      bytes: capture.bytes,
      sha256: capture.hash.digest('hex'),
    };
  }
  return {
    status: result.headers?.status ?? null,
    acceptRanges: result.headers?.acceptRanges ?? null,
    contentLength: result.headers?.contentLength ?? null,
    contentRange: result.headers?.contentRange ?? null,
    contentType: result.headers?.contentType ?? null,
    bytes: boundBytes,
    bytesSent: result.bytesSent ?? null,
    sha256: hash.digest('hex'),
    chunks: result.chunks,
    terminalType: result.terminalType,
    outcome: result.outcome,
    errorCode: result.errorCode,
    captures: captureResults,
    elapsedMs: Date.now() - startedAt,
  };
}

async function fetchAssetInEditor(cdp, context, url, range = null) {
  return evaluateInContext(
    cdp,
    context,
    `(async () => {
      const response = await fetch(${JSON.stringify(url)}, {
        cache: 'no-store',
        ${range ? `headers: { Range: ${JSON.stringify(range)} },` : ''}
      });
      const bytes = new Uint8Array(await response.arrayBuffer());
      const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
      const contentLength = response.headers.get('content-length');
      return {
        status: response.status,
        contentType: response.headers.get('content-type'),
        contentLength: contentLength === null ? null : Number(contentLength),
        contentRange: response.headers.get('content-range'),
        acceptRanges: response.headers.get('accept-ranges'),
        bytes: bytes.byteLength,
        sha256: Array.from(digest, (value) => value.toString(16).padStart(2, '0')).join(''),
        terminalType: 'END',
      };
    })()`,
  );
}

async function compileWasmInEditor(cdp, context, url) {
  return evaluateInContext(
    cdp,
    context,
    `(async () => {
      const response = await fetch(${JSON.stringify(url)}, { cache: 'no-store' });
      const status = response.status;
      const contentType = response.headers.get('content-type');
      await WebAssembly.compileStreaming(Promise.resolve(response));
      return { compiled: true, status, contentType };
    })()`,
  );
}

async function cancelRangeInEditor(cdp, context, url, totalBytes) {
  return evaluateInContext(
    cdp,
    context,
    `(async () => {
      const controller = new AbortController();
      const startedAt = performance.now();
      const response = await fetch(${JSON.stringify(url)}, {
        cache: 'no-store',
        headers: { Range: ${JSON.stringify(`bytes=0-${totalBytes - 1}`)} },
        signal: controller.signal,
      });
      if (!response.body) throw new Error('Editor Broker cancellation response has no body');
      const reader = response.body.getReader();
      const first = await reader.read();
      controller.abort(new DOMException('Acceptance cancellation', 'AbortError'));
      let errorName = null;
      try {
        await reader.read();
      } catch (error) {
        errorName = error?.name ?? 'Error';
      }
      return {
        status: response.status,
        contentLength: Number(response.headers.get('content-length')),
        contentRange: response.headers.get('content-range'),
        acceptRanges: response.headers.get('accept-ranges'),
        firstBytes: first.value?.byteLength ?? 0,
        firstDone: first.done,
        aborted: controller.signal.aborted,
        errorName,
        elapsedMs: performance.now() - startedAt,
      };
    })()`,
  );
}

async function waitForCancelledReadDrain(cdp, options, identity, hostOrigin) {
  const startedAt = Date.now();
  let last = null;
  while (Date.now() - startedAt <= RECOVERY_TIMEOUT_MS) {
    const snapshots = await collectBrokerSnapshots(cdp, options.url);
    const canonical = snapshots.find((snapshot) => snapshot.role === 'canonical-service-worker')?.canonical ?? null;
    const relay =
      snapshots.find(
        (snapshot) =>
          snapshot.role === 'canonical-relay-frame' &&
          snapshot.releaseId === identity.releaseId &&
          snapshot.sessionId === identity.sessionId,
      )?.relay ?? null;
    const editor =
      snapshots.find((snapshot) => {
        if (snapshot.role !== 'editor-service-worker') return false;
        try {
          return new URL(snapshot.targetUrl).origin === hostOrigin;
        } catch {
          return false;
        }
      })?.editor ?? null;
    last = { canonical, relay, editor };
    if (
      canonical?.activeReads === 0 &&
      canonical?.reservedBytes === 0 &&
      (!relay || (relay.activeRequests === 0 && relay.reservedBytes === 0)) &&
      editor?.activeRequests === 0 &&
      editor?.reservedBytes === 0
    ) {
      return { elapsedMs: Date.now() - startedAt, drained: true, ...last };
    }
    await sleep(100);
  }
  return { elapsedMs: Date.now() - startedAt, drained: false, ...last };
}

async function collectDocumentTargetSnapshots(cdp, target) {
  const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: target.targetId, flatten: true });
  const contexts = [];
  cdp.on('Runtime.executionContextCreated', (event, eventSessionId) => {
    if (eventSessionId === sessionId && event.context?.auxData?.isDefault) contexts.push(event.context);
  });
  try {
    await cdp.send('Runtime.disable', {}, sessionId).catch(() => {});
    await cdp.send('Runtime.enable', {}, sessionId);
    await sleep(50);
    const snapshots = [];
    for (const context of contexts) {
      const result = await cdp
        .send(
          'Runtime.evaluate',
          {
            expression: `globalThis.__ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__?.() ?? null`,
            contextId: context.id,
            returnByValue: true,
          },
          sessionId,
        )
        .catch(() => null);
      if (result?.result?.value) {
        snapshots.push({
          ...result.result.value,
          targetUrl: target.url,
          targetId: target.targetId,
          frameId: context.auxData?.frameId ?? null,
        });
      }
    }
    return snapshots;
  } finally {
    await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
  }
}

async function collectBrokerSnapshots(cdp, canonicalOrigin) {
  const { targetInfos } = await cdp.send('Target.getTargets');
  const canonical = new URL(canonicalOrigin);
  const isScopedTarget = (target) => {
    let url;
    try {
      url = new URL(target.url);
    } catch {
      return false;
    }
    return (
      url.origin === canonical.origin ||
      url.hostname.endsWith('.getpi.work') ||
      url.hostname.endsWith('.office.localhost')
    );
  };
  const metricTargets = targetInfos.filter(
    (target) =>
      isScopedTarget(target) &&
      ((target.type === 'service_worker' &&
        (target.url.endsWith('/sw.js') || target.url.endsWith('/document_editor_service_worker.js'))) ||
        target.type === 'page' ||
        target.type === 'iframe'),
  );
  const snapshots = [];
  for (const target of metricTargets) {
    if (target.type === 'service_worker') {
      const metrics = await evaluateTarget(
        cdp,
        target,
        `globalThis.__ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__?.() ?? null`,
      );
      if (metrics) snapshots.push({ ...metrics, targetUrl: target.url, targetId: target.targetId });
    } else {
      snapshots.push(...(await collectDocumentTargetSnapshots(cdp, target).catch(() => [])));
    }
  }
  const unique = new Map();
  for (const snapshot of snapshots) {
    const key =
      snapshot.role === 'canonical-relay-frame'
        ? `${snapshot.role}:${snapshot.targetId}:${snapshot.frameId ?? snapshot.targetUrl}`
        : `${snapshot.role}:${snapshot.targetId}`;
    unique.set(key, snapshot);
  }
  return [...unique.values()];
}

function burstUrl(options, releaseId) {
  const url = new URL('/burst-e2e.html', options.url);
  url.searchParams.set('releaseId', releaseId);
  if (options.hostUrl) url.searchParams.set('hostUrl', options.hostUrl);
  return url.href;
}

async function openBurstTarget(cdp, options, releaseId) {
  const url = burstUrl(options, releaseId);
  const created = await cdp.send('Target.createTarget', { url });
  const target = { targetId: created.targetId, url, createdByScript: true };
  const sessionId = await attachToTarget(cdp, target.targetId);
  await waitForExpression(
    cdp,
    sessionId,
    `Boolean(window.__ONLYOFFICE_BURST_E2E__)`,
    options.timeoutMs,
    'the three-origin burst harness',
  );
  return { target, sessionId };
}

async function startBurst(cdp, sessionId, count, timeoutMs) {
  await evaluate(
    cdp,
    sessionId,
    `window.__ONLYOFFICE_BURST_E2E__.start({ count: ${count}, intervalMs: 0, activationBudget: 1 })`,
  );
  await waitForExpression(
    cdp,
    sessionId,
    `window.__ONLYOFFICE_BURST_E2E__?.getStatus().done === true`,
    timeoutMs,
    `${count} Broker-backed editor(s)`,
  );
  return evaluate(cdp, sessionId, `window.__ONLYOFFICE_BURST_E2E__.getStatus()`);
}

async function currentBurstEditorIdentity(cdp, sessionId, releaseId) {
  const frameSrc = await evaluate(
    cdp,
    sessionId,
    `document.querySelector('#burst-grid iframe.office-editor-host-frame')?.src ?? null`,
  );
  if (typeof frameSrc !== 'string' || !frameSrc) throw new Error('Burst editor Host iframe identity is unavailable');
  const url = new URL(frameSrc);
  const releasePath = `/r/${encodeURIComponent(releaseId)}/office-host.html`;
  const fragment = new URLSearchParams(url.hash.startsWith('#') ? url.hash.slice(1) : url.hash);
  const editorSessionId = fragment.get('sessionId');
  if (url.pathname !== releasePath || !editorSessionId) {
    throw new Error('Burst editor Host iframe has a mismatched release or session identity');
  }
  return {
    hostOrigin: url.origin,
    hostPath: url.pathname,
    sessionId: editorSessionId,
    releaseId,
  };
}

async function closeBurst(cdp, sessionId, timeoutMs) {
  await evaluate(cdp, sessionId, `window.__ONLYOFFICE_BURST_E2E__.closeAll()`);
  await waitForExpression(
    cdp,
    sessionId,
    `window.__ONLYOFFICE_BURST_E2E__?.getStatus().outerFrames === 0`,
    timeoutMs,
    'all burst editors to close',
  );
}

async function waitForConcurrencyDrain(cdp, options, openSnapshots) {
  const startedAt = Date.now();
  const deadline = startedAt + RECOVERY_TIMEOUT_MS;
  const openCanonical = openSnapshots.find((snapshot) => snapshot.role === 'canonical-service-worker')?.canonical;
  const openEditors = openSnapshots.filter((snapshot) => snapshot.role === 'editor-service-worker' && snapshot.editor);
  let snapshots = [];
  let finalCanonical = null;
  let finalEditors = [];
  do {
    snapshots = await collectBrokerSnapshots(cdp, options.url);
    const canonicalSnapshot = snapshots.find((snapshot) => snapshot.role === 'canonical-service-worker')?.canonical;
    finalCanonical =
      canonicalSnapshot ??
      (openCanonical ? { ...openCanonical, activeReads: 0, reservedBytes: 0, lifecycleStatus: 'terminated' } : null);
    finalEditors = openEditors.map((openSnapshot) => {
      const current = snapshots.find(
        (snapshot) =>
          snapshot.role === 'editor-service-worker' && snapshot.targetUrl === openSnapshot.targetUrl && snapshot.editor,
      )?.editor;
      return (
        current ?? {
          ...openSnapshot.editor,
          connectionStatus: 'terminated',
          activeRequests: 0,
          reservedBytes: 0,
        }
      );
    });
    if (
      finalCanonical?.activeReads === 0 &&
      finalCanonical?.reservedBytes === 0 &&
      finalEditors.length === REQUIRED_ORIGINS &&
      finalEditors.every(
        (metric) =>
          ['disconnected', 'terminated'].includes(metric.connectionStatus) &&
          metric.activeRequests === 0 &&
          metric.reservedBytes === 0,
      )
    ) {
      break;
    }
    await sleep(100);
  } while (Date.now() < deadline);
  return { elapsedMs: Date.now() - startedAt, snapshots, finalCanonical, finalEditors };
}

async function runConcurrency(cdp, options, releaseId) {
  const { target, sessionId } = await openBurstTarget(cdp, options, releaseId);
  try {
    const burstStatus = await startBurst(cdp, sessionId, REQUIRED_ORIGINS, options.timeoutMs);
    const openSnapshots = await collectBrokerSnapshots(cdp, options.url);
    await closeBurst(cdp, sessionId, options.timeoutMs);
    const drain = await waitForConcurrencyDrain(cdp, options, openSnapshots);
    return {
      burstStatus,
      openSnapshots,
      finalSnapshots: drain.snapshots,
      drain,
      policy: analyzeBrokerMetrics({
        snapshots: openSnapshots,
        finalCanonicalSnapshot: drain.finalCanonical,
        finalEditorSnapshots: drain.finalEditors,
        burstStatus,
      }),
    };
  } finally {
    await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
    if (!options.keepTargets) await cdp.send('Target.closeTarget', { targetId: target.targetId }).catch(() => {});
  }
}

async function probeCapabilitySecurity(cdp, hostContext) {
  return evaluateInContext(
    cdp,
    hostContext,
    `(async () => {
      const protocol = ${JSON.stringify(BROKER_PROTOCOL)};
      const existing = document.querySelector('iframe[src*="/resource-broker.html?"]');
      if (!(existing instanceof HTMLIFrameElement)) throw new Error('Existing Resource Broker frame is unavailable');
      const brokerOrigin = new URL(existing.src).origin;
      const connect = (frame, capability) => {
        const channel = new MessageChannel();
        const response = new Promise((resolve) => {
          const timer = setTimeout(() => {
            channel.port1.close();
            resolve({ type: 'TIMEOUT' });
          }, 5_000);
          channel.port1.onmessage = (event) => {
            clearTimeout(timer);
            resolve(event.data);
          };
          channel.port1.start();
        });
        frame.contentWindow.postMessage(
          { protocol, type: 'CONNECT', capability },
          brokerOrigin,
          [channel.port2],
        );
        return { response, port: channel.port1 };
      };
      const makeFrame = async (label) => {
        const source = new URL(existing.src);
        source.searchParams.set('sessionId', 'accept-' + label + '-' + crypto.randomUUID());
        const frame = document.createElement('iframe');
        frame.hidden = true;
        frame.sandbox.add('allow-same-origin', 'allow-scripts');
        const challenge = new Promise((resolve, reject) => {
          const timer = setTimeout(() => reject(new Error('Capability challenge timed out')), 10_000);
          const listener = (event) => {
            if (
              event.source !== frame.contentWindow ||
              event.origin !== brokerOrigin ||
              event.data?.protocol !== protocol ||
              event.data?.type !== 'CHALLENGE'
            ) {
              return;
            }
            clearTimeout(timer);
            removeEventListener('message', listener);
            resolve(event.data.capability);
          };
          addEventListener('message', listener);
        });
        frame.src = source.href;
        document.body.appendChild(frame);
        return { frame, capability: await challenge };
      };
      const claim = (capability) => ({
        token: capability.token,
        parentOrigin: capability.parentOrigin,
        editorOrigin: capability.editorOrigin,
        releaseId: capability.releaseId,
        sessionId: capability.sessionId,
      });

      const forgedFrame = await makeFrame('forged');
      const forgedClaim = claim(forgedFrame.capability);
      forgedClaim.token = (forgedClaim.token.startsWith('0') ? '1' : '0') + forgedClaim.token.slice(1);
      const forgedConnection = connect(forgedFrame.frame, forgedClaim);
      const forgedResponse = await forgedConnection.response;
      forgedConnection.port.close();
      forgedFrame.frame.remove();

      const replayFrame = await makeFrame('replay');
      const validClaim = claim(replayFrame.capability);
      const firstConnection = connect(replayFrame.frame, validClaim);
      const firstResponse = await firstConnection.response;
      const secondConnection = connect(replayFrame.frame, validClaim);
      const secondResponse = await secondConnection.response;
      firstConnection.port.close();
      secondConnection.port.close();
      replayFrame.frame.remove();

      return {
        forgedCapability: {
          outcome: forgedResponse?.type ?? 'INVALID',
          errorCode: forgedResponse?.code ?? null,
          connected: forgedResponse?.type === 'CONNECTED',
        },
        replayCapability: {
          firstOutcome: firstResponse?.type ?? 'INVALID',
          secondOutcome: secondResponse?.type ?? 'INVALID',
          secondErrorCode: secondResponse?.code ?? null,
        },
      };
    })()`,
  );
}

async function probeMaliciousOrigin(cdp, brokerFrameUrl) {
  const created = await cdp.send('Target.createTarget', { url: 'about:blank' });
  const sessionId = await attachToTarget(cdp, created.targetId);
  try {
    return await evaluate(
      cdp,
      sessionId,
      `(async () => {
        let challengeCount = 0;
        let connectedCount = 0;
        const listener = (event) => {
          if (event.data?.protocol !== ${JSON.stringify(BROKER_PROTOCOL)}) return;
          if (event.data.type === 'CHALLENGE') challengeCount += 1;
          if (event.data.type === 'CONNECTED') connectedCount += 1;
        };
        addEventListener('message', listener);
        const frame = document.createElement('iframe');
        frame.src = ${JSON.stringify(brokerFrameUrl)};
        document.body.appendChild(frame);
        await new Promise((resolve) => setTimeout(resolve, 1_500));
        frame.remove();
        removeEventListener('message', listener);
        return { origin: 'opaque', challengeCount, connectedCount };
      })()`,
    );
  } finally {
    await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
    await cdp.send('Target.closeTarget', { targetId: created.targetId }).catch(() => {});
  }
}

function brokerIdentityFromContext(context, releaseId) {
  const url = new URL(context.url);
  const sessionId = url.searchParams.get('sessionId');
  if (url.searchParams.get('releaseId') !== releaseId || !sessionId) {
    throw new Error('Resource Broker execution context has a mismatched release identity');
  }
  return { releaseId, sessionId };
}

async function runIntegrity(cdp, options, releaseManifest, hostContext, brokerContext) {
  const identity = brokerIdentityFromContext(brokerContext, releaseManifest.releaseId);
  const fullAsset = releaseManifest.assets.find(
    (asset) => asset.bytes > 0 && asset.bytes <= BROKER_WINDOW_BYTES && !isEditorShellAssetPath(asset.path),
  );
  if (!fullAsset) {
    throw new Error('Release has no <=256 KiB non-shell asset for bounded end-to-end Broker 200 sample');
  }
  const rangeAsset =
    releaseManifest.assets.find(
      (asset) =>
        asset.bytes >= 128 &&
        !isEditorShellAssetPath(asset.path) &&
        asset.fastCdcBoundaries.some((offset) => offset > 1 && offset < asset.bytes - 1),
    ) || releaseManifest.assets.find((asset) => asset.bytes >= 128 && !isEditorShellAssetPath(asset.path));
  if (!rangeAsset) throw new Error('Release has no non-shell asset for a nonzero-offset 206 sample');
  const fastCdcBoundary = rangeAsset.fastCdcBoundaries.find((offset) => offset > 1 && offset < rangeAsset.bytes - 1);
  const rangeStart = fastCdcBoundary
    ? Math.max(1, fastCdcBoundary - 32)
    : Math.max(1, Math.floor((rangeAsset.bytes - 64) / 2));
  const rangeEnd = Math.min(rangeAsset.bytes - 1, rangeStart + 63);
  const cancelAsset = releaseManifest.assets.find(
    (asset) => asset.bytes > BROKER_WINDOW_BYTES && !isEditorShellAssetPath(asset.path),
  );
  if (!cancelAsset) throw new Error('Release has no non-shell asset large enough for a meaningful cancellation probe');
  const assets = [];
  for (let index = 0; index < releaseManifest.assets.length; index += 1) {
    const asset = releaseManifest.assets[index];
    const captureRanges = [];
    if (asset.path === rangeAsset.path) {
      captureRanges.push({ id: 'range-expected', start: rangeStart, end: rangeEnd });
    }
    if (asset.path === cancelAsset.path) {
      captureRanges.push({ id: 'first-byte-expected', start: 0, end: 0 });
    }
    const result = await streamBrokerRead(
      cdp,
      brokerContext,
      {
        protocol: BROKER_PROTOCOL,
        type: 'READ',
        id: `all-${index}`,
        ...identity,
        path: asset.path,
        windowBytes: BROKER_WINDOW_BYTES,
      },
      { timeoutMs: options.timeoutMs, captureRanges },
    );
    assets.push({
      path: asset.path,
      expectedBytes: asset.bytes,
      expectedSha256: asset.sha256,
      expectedMime: asset.mime,
      ...result,
    });
    if ((index + 1) % 25 === 0 || index + 1 === releaseManifest.assets.length) {
      console.log(`Broker release verification: ${index + 1}/${releaseManifest.assets.length} assets`);
    }
  }

  const full = {
    ...(await fetchAssetInEditor(
      cdp,
      hostContext,
      releaseAssetUrl(hostContext.url, releaseManifest.releaseId, fullAsset.path),
    )),
    expectedBytes: fullAsset.bytes,
    expectedSha256: fullAsset.sha256,
    expectedMime: fullAsset.mime,
    path: fullAsset.path,
  };
  const rangeExpected = assets.find((asset) => asset.path === rangeAsset.path)?.captures?.['range-expected'];
  if (!rangeExpected) throw new Error('Full Broker stream did not capture the independent expected Range bytes');
  const directRange = await streamBrokerRead(cdp, brokerContext, {
    protocol: BROKER_PROTOCOL,
    type: 'READ',
    id: 'sample-direct-range',
    ...identity,
    path: rangeAsset.path,
    range: { kind: 'closed', start: rangeStart, end: rangeEnd },
    windowBytes: BROKER_WINDOW_BYTES,
  });
  const range = {
    ...(await fetchAssetInEditor(
      cdp,
      hostContext,
      releaseAssetUrl(hostContext.url, releaseManifest.releaseId, rangeAsset.path),
      `bytes=${rangeStart}-${rangeEnd}`,
    )),
    expectedContentRange: `bytes ${rangeStart}-${rangeEnd}/${rangeAsset.bytes}`,
    expectedSha256: rangeExpected.sha256,
    directSha256: directRange.sha256,
    expectedMime: rangeAsset.mime,
    expectedSource: 'full-broker-stream-capture',
    start: rangeStart,
    end: rangeEnd,
    crossedFastCdcBoundary: fastCdcBoundary ? rangeStart < fastCdcBoundary && rangeEnd >= fastCdcBoundary : false,
    path: rangeAsset.path,
  };
  const unsatisfied = {
    ...(await streamBrokerRead(cdp, brokerContext, {
      protocol: BROKER_PROTOCOL,
      type: 'READ',
      id: 'sample-unsatisfied',
      ...identity,
      path: rangeAsset.path,
      range: { kind: 'open', start: rangeAsset.bytes },
      windowBytes: BROKER_WINDOW_BYTES,
    })),
    expectedContentRange: `bytes */${rangeAsset.bytes}`,
  };
  const cancellation = {
    ...(await streamBrokerRead(
      cdp,
      brokerContext,
      {
        protocol: BROKER_PROTOCOL,
        type: 'READ',
        id: 'sample-cancel',
        ...identity,
        path: cancelAsset.path,
        range: { kind: 'closed', start: 0, end: cancelAsset.bytes - 1 },
        windowBytes: BROKER_WINDOW_BYTES,
      },
      { cancelAfterFirstChunk: true },
    )),
    expectedBytes: cancelAsset.bytes,
  };
  cancellation.editor = {
    ...(await cancelRangeInEditor(
      cdp,
      hostContext,
      releaseAssetUrl(hostContext.url, releaseManifest.releaseId, cancelAsset.path),
      cancelAsset.bytes,
    )),
    expectedContentRange: `bytes 0-${cancelAsset.bytes - 1}/${cancelAsset.bytes}`,
    expectedBytes: cancelAsset.bytes,
  };
  cancellation.drain = await waitForCancelledReadDrain(cdp, options, identity, new URL(hostContext.url).origin);
  cancellation.protocolRecovery = {
    ...(await streamBrokerRead(cdp, brokerContext, {
      protocol: BROKER_PROTOCOL,
      type: 'READ',
      id: 'sample-after-cancel',
      ...identity,
      path: cancelAsset.path,
      range: { kind: 'closed', start: 0, end: 0 },
      windowBytes: BROKER_WINDOW_BYTES,
    })),
    expectedContentRange: `bytes 0-0/${cancelAsset.bytes}`,
  };
  cancellation.recovery = {
    ...(await fetchAssetInEditor(
      cdp,
      hostContext,
      releaseAssetUrl(hostContext.url, releaseManifest.releaseId, cancelAsset.path),
      'bytes=0-0',
    )),
    expectedContentRange: `bytes 0-0/${cancelAsset.bytes}`,
    expectedSha256: assets.find((asset) => asset.path === cancelAsset.path)?.captures?.['first-byte-expected']?.sha256,
    expectedSource: 'full-broker-stream-capture',
  };
  const wasmAsset = releaseManifest.assets.find((asset) => asset.path.endsWith('.wasm'));
  if (!wasmAsset) throw new Error('Release manifest has no WASM asset for compileStreaming');
  const wasm = await compileWasmInEditor(
    cdp,
    hostContext,
    releaseAssetUrl(hostContext.url, releaseManifest.releaseId, wasmAsset.path),
  );
  const input = {
    expectedAssetCount: releaseManifest.assets.length,
    assets,
    samples: { full, range, unsatisfied, cancellation, wasm },
  };
  return {
    manifest: {
      url: releaseManifest.url,
      sha256: releaseManifest.manifestSha256,
      releaseId: releaseManifest.releaseId,
      assetCount: releaseManifest.assets.length,
    },
    input,
    policy: analyzeReleaseIntegrity(input),
  };
}

async function runSecurity(cdp, options, releaseManifest, hostContext, brokerContext) {
  const identity = brokerIdentityFromContext(brokerContext, releaseManifest.releaseId);
  const capability = await probeCapabilitySecurity(cdp, hostContext);
  const maliciousOrigin = await probeMaliciousOrigin(cdp, brokerContext.url);
  const validPath = releaseManifest.assets[0]?.path;
  if (!validPath) throw new Error('Release manifest has no asset for security probes');
  const probes = [
    { kind: 'syntax', path: `./${validPath}` },
    { kind: 'syntax', path: `%2e/${validPath}` },
    { kind: 'syntax', path: `nested/%2e%2e/${validPath}` },
    {
      kind: 'syntax',
      path: releaseAssetUrl(options.url, releaseManifest.releaseId, validPath),
    },
    {
      kind: 'manifest-boundary',
      path: `__onlyoffice_content__/sha256/${releaseManifest.assets[0].sha256}`,
    },
  ];
  const pathProbes = [];
  for (let index = 0; index < probes.length; index += 1) {
    const probe = probes[index];
    const result = await streamBrokerRead(
      cdp,
      brokerContext,
      {
        protocol: BROKER_PROTOCOL,
        type: 'READ',
        id: `security-path-${index}`,
        ...identity,
        path: probe.path,
        windowBytes: BROKER_WINDOW_BYTES,
      },
      { timeoutMs: 2_000 },
    );
    pathProbes.push({ kind: probe.kind, probe: probe.path, ...result });
  }
  const sampleAsset = releaseManifest.assets.find((asset) => asset.bytes > 0 && !isEditorShellAssetPath(asset.path));
  if (!sampleAsset) throw new Error('Release has no non-shell asset for oversized Broker Range probe');
  const oversizedRange = {
    ...(await fetchAssetInEditor(
      cdp,
      hostContext,
      releaseAssetUrl(hostContext.url, releaseManifest.releaseId, sampleAsset.path),
      'bytes=9007199254740991-',
    )),
    expectedContentRange: `bytes */${sampleAsset.bytes}`,
  };
  const input = {
    maliciousOrigin,
    ...capability,
    pathProbes,
    oversizedRange,
  };
  return { input, policy: analyzeSecurityProbes(input) };
}

async function runProtocolSections(cdp, options, releaseManifest) {
  const { target, sessionId } = await openBurstTarget(cdp, options, releaseManifest.releaseId);
  let hostContext = null;
  let brokerContext = null;
  try {
    const burstStatus = await startBurst(cdp, sessionId, 1, options.timeoutMs);
    if (burstStatus.ready !== 1 || burstStatus.errors.length) {
      throw new Error(`Protocol probe editor did not become ready: ${burstStatus.errors.join('; ')}`);
    }
    const editorIdentity = await currentBurstEditorIdentity(cdp, sessionId, releaseManifest.releaseId);
    const canonicalOrigin = new URL(options.url).origin;
    hostContext = await findExecutionContext(
      cdp,
      (value) => {
        try {
          const url = new URL(value);
          const fragment = new URLSearchParams(url.hash.startsWith('#') ? url.hash.slice(1) : url.hash);
          return (
            url.origin === editorIdentity.hostOrigin &&
            url.pathname === editorIdentity.hostPath &&
            fragment.get('sessionId') === editorIdentity.sessionId
          );
        } catch {
          return false;
        }
      },
      'Office Host',
      options.timeoutMs,
    );
    brokerContext = await findExecutionContext(
      cdp,
      (value) => {
        try {
          const url = new URL(value);
          return (
            url.origin === canonicalOrigin &&
            url.pathname === '/resource-broker.html' &&
            url.searchParams.get('releaseId') === editorIdentity.releaseId &&
            url.searchParams.get('sessionId') === editorIdentity.sessionId
          );
        } catch {
          return false;
        }
      },
      'canonical Resource Broker',
      options.timeoutMs,
    );
    const result = { burstStatus, editorIdentity };
    if (options.sections.has('integrity')) {
      result.integrity = await runIntegrity(cdp, options, releaseManifest, hostContext, brokerContext);
    }
    if (options.sections.has('security')) {
      result.security = await runSecurity(cdp, options, releaseManifest, hostContext, brokerContext);
    }
    return result;
  } finally {
    await detachExecutionContext(cdp, brokerContext);
    await detachExecutionContext(cdp, hostContext);
    await closeBurst(cdp, sessionId, options.timeoutMs).catch(() => {});
    await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
    if (!options.keepTargets) await cdp.send('Target.closeTarget', { targetId: target.targetId }).catch(() => {});
  }
}

async function stopMatchingServiceWorker(cdp, matcher, label) {
  const { targetInfos } = await cdp.send('Target.getTargets');
  const target = targetInfos.find(
    (candidate) => candidate.type === 'service_worker' && matcher(new URL(candidate.url)),
  );
  if (!target) throw new Error(`Unable to find the active ${label} Service Worker`);
  await cdp.send('Target.closeTarget', { targetId: target.targetId });
  return { targetId: target.targetId, scriptURL: target.url, action: 'Target.closeTarget' };
}

async function resourceBrokerFrameTarget(cdp, identity) {
  const { targetInfos } = await cdp.send('Target.getTargets');
  const target = targetInfos.find((candidate) => {
    if (candidate.type !== 'iframe') return false;
    try {
      const url = new URL(candidate.url);
      return (
        url.pathname === '/resource-broker.html' &&
        url.searchParams.get('releaseId') === identity.releaseId &&
        url.searchParams.get('sessionId') === identity.sessionId
      );
    } catch {
      return false;
    }
  });
  if (!target) throw new Error(`Canonical Broker iframe target is unavailable for ${identity.sessionId}`);
  return target;
}

async function ensureAttachedSession(cdp, targetId, sessionId) {
  const alive = await cdp
    .send('Runtime.evaluate', { expression: 'true', returnByValue: true }, sessionId)
    .then(() => true)
    .catch(() => false);
  return alive ? sessionId : attachToTarget(cdp, targetId);
}

async function runRecoveryCase(cdp, options, targetId, initialSessionId, name, disrupt) {
  let sessionId = await ensureAttachedSession(cdp, targetId, initialSessionId);
  await closeBurst(cdp, sessionId, options.timeoutMs);
  const initial = await startBurst(cdp, sessionId, 1, options.timeoutMs);
  if (initial.ready !== 1 || initial.errors.length) throw new Error(`${name} setup editor did not become ready`);
  const identity = await currentBurstEditorIdentity(cdp, sessionId, options.releaseId);
  const startedAt = Date.now();
  let disruption = null;
  try {
    disruption = await disrupt(identity, sessionId);
    sessionId = await ensureAttachedSession(cdp, targetId, sessionId);
    await closeBurst(cdp, sessionId, RECOVERY_TIMEOUT_MS);
    const recovered = await startBurst(cdp, sessionId, 1, RECOVERY_TIMEOUT_MS);
    const elapsedMs = Date.now() - startedAt;
    if (recovered.ready !== 1 || recovered.errors.length) {
      return {
        name,
        outcome: 'failed',
        elapsedMs,
        errors: recovered.errors,
        disruption,
      };
    }
    return { name, outcome: 'recovered', elapsedMs, disruption };
  } catch (error) {
    return {
      name,
      outcome: 'failed',
      elapsedMs: Date.now() - startedAt,
      error: error instanceof Error ? error.message : String(error),
      disruption,
    };
  } finally {
    await closeBurst(cdp, sessionId, options.timeoutMs).catch(() => {});
  }
}

async function runRecovery(cdp, options, releaseId) {
  const canonicalOrigin = new URL(options.url).origin;
  const recoveryOptions = { ...options, releaseId };
  const runIsolatedCase = async (name, disrupt) => {
    const { target, sessionId } = await openBurstTarget(cdp, options, releaseId);
    try {
      return await runRecoveryCase(cdp, recoveryOptions, target.targetId, sessionId, name, disrupt);
    } finally {
      await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
      if (!options.keepTargets) await cdp.send('Target.closeTarget', { targetId: target.targetId }).catch(() => {});
    }
  };
  const results = [];
  results.push(
    await runIsolatedCase('broker-iframe', async (identity) => {
      const targetInfo = await resourceBrokerFrameTarget(cdp, identity);
      const relaySession = await attachToTarget(cdp, targetInfo.targetId);
      try {
        await cdp.send(
          'Runtime.evaluate',
          { expression: 'queueMicrotask(() => location.reload()); true', returnByValue: true },
          relaySession,
        );
      } finally {
        await cdp.send('Target.detachFromTarget', { sessionId: relaySession }).catch(() => {});
      }
      return { targetId: targetInfo.targetId, action: 'Page.reload' };
    }),
  );
  results.push(
    await runIsolatedCase('canonical-service-worker', () =>
      stopMatchingServiceWorker(cdp, (url) => url.origin === canonicalOrigin && url.pathname === '/sw.js', 'canonical'),
    ),
  );
  results.push(
    await runIsolatedCase('editor-service-worker', (identity) =>
      stopMatchingServiceWorker(
        cdp,
        (url) =>
          url.origin === identity.hostOrigin &&
          (url.pathname === '/sw.js' || url.pathname === '/document_editor_service_worker.js'),
        'editor',
      ),
    ),
  );
  results.push(
    await runIsolatedCase('message-port', async (identity) => {
      const targetInfo = await resourceBrokerFrameTarget(cdp, identity);
      const relaySession = await attachToTarget(cdp, targetInfo.targetId);
      try {
        await cdp.send(
          'Runtime.evaluate',
          { expression: "dispatchEvent(new Event('pagehide')); true", returnByValue: true },
          relaySession,
        );
      } finally {
        await cdp.send('Target.detachFromTarget', { sessionId: relaySession }).catch(() => {});
      }
      return { targetId: targetInfo.targetId, action: 'dispatch pagehide and close MessagePort' };
    }),
  );
  results.push(
    await runIsolatedCase('freeze-resume', async (_identity, sessionId) => {
      await cdp.send('Page.setWebLifecycleState', { state: 'frozen' }, sessionId);
      await sleep(500);
      await cdp.send('Page.setWebLifecycleState', { state: 'active' }, sessionId);
      return { action: 'frozen->active', frozenMs: 500 };
    }),
  );
  return { results, policy: analyzeRecoveryResults(results) };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const baseline = options.sections.has('lifecycle') ? resolveBaseline(options) : null;
  mkdirSync(dirname(options.output), { recursive: true });
  const report = {
    schemaVersion: 1,
    startedAt: new Date().toISOString(),
    completedAt: null,
    pass: false,
    fixedGates: {
      lifecycleIterations: REQUIRED_ITERATIONS,
      concurrentOrigins: REQUIRED_ORIGINS,
      recoveryTimeoutMs: RECOVERY_TIMEOUT_MS,
      brokerBufferHighWaterBytes: 64 * 1024 * 1024,
      startupExtraP95: 'max(300ms, baselineP95*10%)',
      brokerReadWindowBytes: BROKER_WINDOW_BYTES,
      releaseIntegrity: 'every manifest asset exact SHA-256/MIME/length',
      trafficEvidence: 'zero content segment/object Worker/cache/R2/byte/status counter delta',
      httpCache: 'cleared before the traffic counter baseline',
    },
    options: {
      browserWs: options.browserWs,
      url: options.url,
      hostUrl: options.hostUrl,
      profileDir: options.profileDir,
      sections: [...options.sections],
      trafficEvidenceUrl: options.trafficEvidenceUrl,
    },
    baseline,
    preflight: null,
    evidence: {},
    result: null,
    error: null,
  };
  writeFileSync(options.output, JSON.stringify(report, null, 2));
  console.log('Fixed gates:', JSON.stringify(report.fixedGates));
  console.log(`Auditable report: ${options.output}`);

  const cdp = new CdpClient(options.browserWs);
  try {
    await cdp.connect();
    report.preflight = await preflight(cdp, options);
    const releaseManifest = await loadReleaseManifest(options, report.preflight.releaseId);
    report.evidence.releaseManifest = {
      url: releaseManifest.url,
      sha256: releaseManifest.manifestSha256,
      releaseId: releaseManifest.releaseId,
      assetCount: releaseManifest.assets.length,
      totalBytes: releaseManifest.assets.reduce((sum, asset) => sum + asset.bytes, 0),
    };
    report.evidence.httpCacheClear = await clearBrowserHttpCache(cdp, report.preflight.targetId);
    const trafficBefore = await fetchTrafficEvidence(options);
    const sections = {};
    if (options.sections.has('lifecycle')) {
      const lifecycle = await runLifecycle(options, baseline);
      report.evidence.lifecycle = lifecycle;
      sections.lifecycle = lifecycle.policy;
      sections.startup = lifecycle.startup;
    }
    if (options.sections.has('concurrency')) {
      const concurrency = await runConcurrency(cdp, options, report.preflight.releaseId);
      report.evidence.concurrency = concurrency;
      sections.concurrency = concurrency.policy;
    }
    if (options.sections.has('recovery')) {
      const recovery = await runRecovery(cdp, options, report.preflight.releaseId);
      report.evidence.recovery = recovery;
      sections.recovery = recovery.policy;
    }
    if (options.sections.has('integrity') || options.sections.has('security')) {
      const protocol = await runProtocolSections(cdp, options, releaseManifest);
      report.evidence.protocol = protocol;
      if (protocol.integrity) sections.integrity = protocol.integrity.policy;
      if (protocol.security) sections.security = protocol.security.policy;
    }
    const trafficAfter = await fetchTrafficEvidence(options);
    const traffic = {
      evidenceUrl: options.trafficEvidenceUrl,
      before: trafficBefore,
      after: trafficAfter,
    };
    report.evidence.traffic = traffic;
    sections.traffic = analyzeTrafficEvidence(traffic);
    report.result = combineAcceptanceSections(sections);
    report.pass = report.result.pass;
    if (!report.pass) process.exitCode = 1;
  } catch (error) {
    report.error = {
      name: error instanceof Error ? error.name : 'Error',
      message: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : null,
    };
    process.exitCode = 1;
  } finally {
    report.completedAt = new Date().toISOString();
    writeFileSync(options.output, JSON.stringify(report, null, 2));
    if (report.preflight?.targetCreatedByRunner && !options.keepTargets) {
      await cdp.send('Target.closeTarget', { targetId: report.preflight.targetId }).catch(() => {});
    }
    cdp.close();
    console.log(`Canonical Broker acceptance: ${report.pass ? 'PASS' : 'FAIL'}`);
    console.log(`Report: ${options.output}`);
  }
}

main();
