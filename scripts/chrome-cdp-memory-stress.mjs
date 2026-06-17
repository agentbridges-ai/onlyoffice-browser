#!/usr/bin/env node

import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { basename, dirname, extname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_BROWSER_WS = 'ws://localhost:9222/devtools/browser';
const DEFAULT_URL = 'http://127.0.0.1:5173/';
const DEFAULT_ITERATIONS = 100;
const DEFAULT_STAY_MS = 2_000;
const DEFAULT_HOME_DWELL_MS = 50;
const DEFAULT_TIMEOUT_MS = 180_000;
const RESET_STABILITY_MS = 1_000;
const DEFAULT_FORMATS = ['docx', 'xlsx', 'pptx'];
const DEFAULT_MODES = ['edit'];
const DEFAULT_SEED = 930_140_100;
const OFFICE_EDITOR_MODES = ['edit', 'readonly', 'preview'];
const FORMAT_SELECTORS = {
  docx: '#new-word-button',
  xlsx: '#new-excel-button',
  pptx: '#new-pptx-button',
};
const FORMAT_LABELS = {
  docx: 'Word',
  xlsx: 'Excel',
  pptx: 'PowerPoint',
};

class CdpClient {
  constructor(wsUrl) {
    this.wsUrl = wsUrl;
    this.nextId = 1;
    this.pending = new Map();
    this.listeners = new Map();
    this.ws = null;
  }

  async connect() {
    this.ws = new WebSocket(this.wsUrl);
    this.ws.addEventListener('message', (event) => {
      const message = JSON.parse(event.data);
      if (message.id) {
        const pending = this.pending.get(message.id);
        if (!pending) return;
        this.pending.delete(message.id);
        if (message.error) {
          pending.reject(new Error(`${message.error.message}${message.error.data ? `: ${message.error.data}` : ''}`));
        } else {
          pending.resolve(message.result ?? {});
        }
        return;
      }

      const callbacks = this.listeners.get(message.method);
      if (!callbacks) return;
      for (const callback of callbacks) {
        callback(message.params ?? {}, message.sessionId);
      }
    });

    await new Promise((resolve, reject) => {
      this.ws.addEventListener('open', resolve, { once: true });
      this.ws.addEventListener('error', reject, { once: true });
    });
  }

  on(method, callback) {
    const callbacks = this.listeners.get(method) ?? [];
    callbacks.push(callback);
    this.listeners.set(method, callbacks);
  }

  send(method, params = {}, sessionId = undefined) {
    if (!this.ws || this.ws.readyState !== WebSocket.OPEN) {
      return Promise.reject(new Error('CDP websocket is not open'));
    }

    const id = this.nextId;
    this.nextId += 1;
    const message = { id, method, params };
    if (sessionId) {
      message.sessionId = sessionId;
    }

    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify(message));
    });
  }

  close() {
    this.ws?.close();
  }
}

function parseArgs(argv) {
  const options = {
    browserWs: DEFAULT_BROWSER_WS,
    url: DEFAULT_URL,
    iterations: DEFAULT_ITERATIONS,
    stayMs: DEFAULT_STAY_MS,
    homeDwellMs: DEFAULT_HOME_DWELL_MS,
    timeoutMs: DEFAULT_TIMEOUT_MS,
    formats: DEFAULT_FORMATS,
    modes: DEFAULT_MODES,
    seed: DEFAULT_SEED,
    batchSize: 1,
    gcPasses: 2,
    finalGcPasses: 5,
    finalSettleMs: 0,
    openedSampleInterval: 10,
    chromePid: null,
    createTargetIfMissing: true,
    keepTarget: false,
    closeMode: 'back',
    hardResetOnClose: false,
    requireRealFiles: false,
    filePaths: [],
  };

  for (const arg of argv) {
    if (arg.startsWith('--browser-ws=')) options.browserWs = arg.slice('--browser-ws='.length);
    else if (arg.startsWith('--url=')) options.url = arg.slice('--url='.length);
    else if (arg.startsWith('--iterations=')) options.iterations = Number(arg.slice('--iterations='.length));
    else if (arg.startsWith('--stay-ms=')) options.stayMs = Number(arg.slice('--stay-ms='.length));
    else if (arg.startsWith('--home-dwell-ms=')) options.homeDwellMs = Number(arg.slice('--home-dwell-ms='.length));
    else if (arg.startsWith('--timeout-ms=')) options.timeoutMs = Number(arg.slice('--timeout-ms='.length));
    else if (arg.startsWith('--seed=')) options.seed = Number(arg.slice('--seed='.length));
    else if (arg.startsWith('--batch-size=')) options.batchSize = Number(arg.slice('--batch-size='.length));
    else if (arg.startsWith('--gc-passes=')) options.gcPasses = Number(arg.slice('--gc-passes='.length));
    else if (arg.startsWith('--final-gc-passes=')) options.finalGcPasses = Number(arg.slice('--final-gc-passes='.length));
    else if (arg.startsWith('--final-settle-ms=')) options.finalSettleMs = Number(arg.slice('--final-settle-ms='.length));
    else if (arg.startsWith('--opened-sample-interval='))
      options.openedSampleInterval = Number(arg.slice('--opened-sample-interval='.length));
    else if (arg.startsWith('--chrome-pid=')) options.chromePid = Number(arg.slice('--chrome-pid='.length));
    else if (arg.startsWith('--close-mode=')) options.closeMode = arg.slice('--close-mode='.length);
    else if (arg === '--hard-reset-on-close') options.hardResetOnClose = true;
    else if (arg === '--require-real-files') options.requireRealFiles = true;
    else if (arg.startsWith('--file=')) options.filePaths.push(arg.slice('--file='.length));
    else if (arg === '--require-existing-tab') options.createTargetIfMissing = false;
    else if (arg === '--keep-target') options.keepTarget = true;
    else if (arg.startsWith('--formats='))
      options.formats = arg
        .slice('--formats='.length)
        .split(',')
        .map((format) => format.trim().replace(/^\./, '').toLowerCase())
        .filter(Boolean);
    else if (arg.startsWith('--modes='))
      options.modes = arg
        .slice('--modes='.length)
        .split(',')
        .map((mode) => mode.trim().toLowerCase())
        .filter(Boolean);
    else if (arg === '--help' || arg === '-h') {
      printHelp();
      process.exit(0);
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }

  if (!Number.isFinite(options.iterations) || options.iterations < 1) {
    throw new Error('--iterations must be a positive number');
  }
  if (!Number.isFinite(options.stayMs) || options.stayMs < 0) {
    throw new Error('--stay-ms must be a non-negative number');
  }
  if (!Number.isFinite(options.homeDwellMs) || options.homeDwellMs < 0) {
    throw new Error('--home-dwell-ms must be a non-negative number');
  }
  if (!Number.isFinite(options.seed)) {
    throw new Error('--seed must be a finite number');
  }
  if (!Number.isInteger(options.batchSize) || options.batchSize < 1) {
    throw new Error('--batch-size must be a positive integer');
  }
  if (!Number.isFinite(options.gcPasses) || options.gcPasses < 0) {
    throw new Error('--gc-passes must be a non-negative number');
  }
  if (!Number.isFinite(options.finalGcPasses) || options.finalGcPasses < 0) {
    throw new Error('--final-gc-passes must be a non-negative number');
  }
  if (!Number.isFinite(options.finalSettleMs) || options.finalSettleMs < 0) {
    throw new Error('--final-settle-ms must be a non-negative number');
  }
  if (!Number.isFinite(options.openedSampleInterval) || options.openedSampleInterval < 0) {
    throw new Error('--opened-sample-interval must be a non-negative number');
  }
  if (!options.formats.length || options.formats.some((format) => !FORMAT_SELECTORS[format])) {
    throw new Error('--formats must contain one or more of: docx,xlsx,pptx');
  }
  if (!options.modes.length || options.modes.some((mode) => !OFFICE_EDITOR_MODES.includes(mode))) {
    throw new Error('--modes must contain one or more of: edit,readonly,preview');
  }
  if (!['back', 'direct', 'app'].includes(options.closeMode)) {
    throw new Error('--close-mode must be back, direct, or app');
  }
  options.filePaths = options.filePaths.map((filePath) => resolve(filePath));
  if (options.requireRealFiles && options.filePaths.length === 0) {
    throw new Error('--require-real-files needs at least one --file=/path/to/real-office-file argument');
  }
  if (options.requireRealFiles && options.batchSize > 1 && options.filePaths.length < options.batchSize) {
    throw new Error('--require-real-files with --batch-size > 1 needs at least batch-size distinct --file arguments');
  }
  for (const filePath of options.filePaths) {
    if (!existsSync(filePath)) {
      throw new Error(`File does not exist: ${filePath}`);
    }
  }

  return options;
}

function printHelp() {
  console.log(`Usage: npm run memory:chrome-cdp -- [options]
       node scripts/chrome-cdp-memory-stress.mjs --browser-ws=ws://localhost:9222/devtools/browser

Options:
  --browser-ws=ws://localhost:9222/devtools/browser
  --url=http://127.0.0.1:5173/       App URL to stress in the current Chrome.
  --iterations=100                    Random create/close cycles in one tab.
  --formats=docx,xlsx,pptx            Comma-separated formats.
  --modes=edit                        Comma-separated open modes: edit,readonly,preview.
  --stay-ms=2000                      Time to keep each editor open.
  --home-dwell-ms=50                  Time to stay on the home screen after it is ready.
  --seed=930140100                    Deterministic random seed.
  --batch-size=1                      Open N documents simultaneously in each cycle.
  --gc-passes=2                       HeapProfiler.collectGarbage passes after each close.
  --final-gc-passes=5                 Extra GC passes before final sample.
  --final-settle-ms=0                 Extra idle time before a settled final sample.
  --opened-sample-interval=10         Sample opened state every N cycles; 0 disables.
  --chrome-pid=14337                  Override browser root PID for RSS sampling.
  --close-mode=back                   Close through browser back, app close control, or direct destroy.
  --hard-reset-on-close               Enable the demo hard-reset toggle before each open.
  --require-real-files                Fail unless at least one --file argument is provided.
  --file=/path/to/document.docx        Add a real local Office file workload. Repeatable.
  --require-existing-tab              Fail if no matching app tab already exists.
  --keep-target                        Keep a Chrome target created by this script open after exit.
`);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function bytesToMb(value) {
  if (!Number.isFinite(value)) return null;
  return value / 1024 / 1024;
}

function formatMb(value) {
  return Number.isFinite(value) ? value.toFixed(1) : 'n/a';
}

function createRandom(seed) {
  let value = seed >>> 0;
  return () => {
    value += 0x6d2b79f5;
    let next = value;
    next = Math.imul(next ^ (next >>> 15), next | 1);
    next ^= next + Math.imul(next ^ (next >>> 7), next | 61);
    return ((next ^ (next >>> 14)) >>> 0) / 4294967296;
  };
}

function pickRandom(random, values) {
  return values[Math.floor(random() * values.length)];
}

function getFileFormat(filePath) {
  return extname(filePath).replace(/^\./, '').toLowerCase() || 'file';
}

function getWorkloadLabel(workload) {
  if (workload.items) {
    return workload.items.map(getWorkloadLabel).join(' + ');
  }
  const modeSuffix = workload.mode ? ` [${workload.mode}]` : '';
  if (workload.filePath) {
    return `${basename(workload.filePath)}${modeSuffix}`;
  }
  return `${FORMAT_LABELS[workload.format] ?? workload.format} .${workload.format}${modeSuffix}`;
}

function getWorkloadFormat(workload) {
  if (workload?.items) {
    return `batch${workload.items.length}`;
  }
  return workload?.format ?? null;
}

function getWorkloadFileName(workload) {
  if (workload?.items) {
    return workload.items.map((item) => (item.filePath ? basename(item.filePath) : item.format)).join(' + ');
  }
  return workload?.filePath ? basename(workload.filePath) : null;
}

function getWorkloadFilePath(workload) {
  if (workload?.items) {
    return workload.items.map((item) => item.filePath ?? item.format).join(' + ');
  }
  return workload?.filePath ?? null;
}

function getWorkloadMode(workload) {
  if (workload?.items) {
    return workload.items.map((item) => item.mode ?? 'edit').join(' + ');
  }
  return workload?.mode ?? null;
}

function getWorkloadReadyWaitMs(workload) {
  return workload?.readyWaitMs ?? null;
}

function assignModeToWorkload(workload, mode) {
  return { ...workload, mode };
}

function shuffleWithRandom(random, values) {
  return [...values].sort(() => random() - 0.5);
}

function pickWorkloadBatch(random, workloads, batchSize, modes) {
  if (batchSize <= 1) {
    return assignModeToWorkload(pickRandom(random, workloads), pickRandom(random, modes));
  }

  const shuffled = shuffleWithRandom(random, workloads);
  const shuffledModes = shuffleWithRandom(random, modes);
  const items = [];
  while (items.length < batchSize) {
    items.push(assignModeToWorkload(shuffled[items.length % shuffled.length], shuffledModes[items.length % shuffledModes.length]));
  }
  return { items, format: `batch${items.length}` };
}

function metricValue(metrics, name) {
  const metric = metrics.find((item) => item.name === name);
  return metric ? metric.value : null;
}

function normalizeUrlPrefix(url) {
  const parsed = new URL(url);
  parsed.hash = '';
  return parsed.href.replace(/\/$/, '');
}

function inferDebugPort(browserWs) {
  try {
    const parsed = new URL(browserWs);
    return Number(parsed.port);
  } catch {
    return null;
  }
}

function findChromeRootPid(debugPort) {
  if (!debugPort) return null;

  try {
    const output = execFileSync('lsof', ['-nP', `-iTCP:${debugPort}`, '-sTCP:LISTEN', '-t'], { encoding: 'utf8' });
    const firstPid = output
      .trim()
      .split(/\s+/)
      .map(Number)
      .find(Number.isFinite);
    return firstPid ?? null;
  } catch {
    return null;
  }
}

function collectDescendantPids(rootPid) {
  if (!rootPid) return [];

  let output = '';
  try {
    output = execFileSync('ps', ['-axo', 'pid=,ppid=,rss=,command='], { encoding: 'utf8' });
  } catch {
    return [];
  }

  const processes = output
    .trim()
    .split('\n')
    .map((line) => {
      const match = line.trim().match(/^(\d+)\s+(\d+)\s+(\d+)\s+(.+)$/);
      if (!match) return null;
      return {
        pid: Number(match[1]),
        ppid: Number(match[2]),
        rssKb: Number(match[3]),
        command: match[4],
      };
    })
    .filter(Boolean);

  const byParent = new Map();
  for (const proc of processes) {
    if (!byParent.has(proc.ppid)) byParent.set(proc.ppid, []);
    byParent.get(proc.ppid).push(proc);
  }

  const result = [];
  const seen = new Set();
  const stack = [rootPid];
  while (stack.length) {
    const pid = stack.pop();
    if (seen.has(pid)) continue;
    seen.add(pid);
    const current = processes.find((proc) => proc.pid === pid);
    if (current) result.push(current);
    for (const child of byParent.get(pid) ?? []) {
      stack.push(child.pid);
    }
  }

  return result;
}

function classifyChromeProcess(proc, rootPid) {
  const command = proc.command.toLowerCase();
  if (proc.pid === rootPid) return 'browser';
  if (command.includes('crashpad')) return 'crashpad';
  if (command.includes('renderer') || command.includes('--type=renderer')) return 'renderer';
  if (command.includes('gpu') || command.includes('--type=gpu-process')) return 'gpu';
  if (command.includes('utility') || command.includes('--type=utility')) return 'utility';
  if (command.includes('zygote') || command.includes('--type=zygote')) return 'zygote';
  if (command.includes('network') || command.includes('--utility-sub-type=network')) return 'network';
  return 'other';
}

function summarizeBrowserProcesses(processes, rootPid) {
  const breakdown = {};
  for (const proc of processes) {
    const type = classifyChromeProcess(proc, rootPid);
    const current = breakdown[type] ?? { count: 0, rssMb: 0 };
    current.count += 1;
    current.rssMb += proc.rssKb / 1024;
    breakdown[type] = current;
  }

  const topProcesses = [...processes]
    .sort((a, b) => b.rssKb - a.rssKb)
    .slice(0, 10)
    .map((proc) => ({
      pid: proc.pid,
      ppid: proc.ppid,
      type: classifyChromeProcess(proc, rootPid),
      rssMb: proc.rssKb / 1024,
      command: proc.command.length > 180 ? `${proc.command.slice(0, 177)}...` : proc.command,
    }));

  return { breakdown, topProcesses };
}

function collectBrowserRssMb(rootPid) {
  const processes = collectDescendantPids(rootPid);
  if (!processes.length) {
    return {
      totalMb: null,
      processCount: 0,
      breakdown: {},
      topProcesses: [],
    };
  }

  const totalKb = processes.reduce((sum, proc) => sum + proc.rssKb, 0);
  const summary = summarizeBrowserProcesses(processes, rootPid);

  return {
    totalMb: totalKb / 1024,
    processCount: processes.length,
    ...summary,
  };
}

async function evaluate(cdp, sessionId, expression, options = {}) {
  const result = await cdp.send(
    'Runtime.evaluate',
    {
      expression,
      awaitPromise: true,
      returnByValue: true,
      userGesture: Boolean(options.userGesture),
    },
    sessionId,
  );

  if (result.exceptionDetails) {
    const description = result.exceptionDetails.exception?.description || result.exceptionDetails.text;
    throw new Error(description);
  }

  return result.result?.value;
}

async function waitForExpression(cdp, sessionId, expression, timeoutMs, label) {
  const started = Date.now();
  let lastError = null;
  while (Date.now() - started < timeoutMs) {
    try {
      if (await evaluate(cdp, sessionId, expression)) return;
    } catch (error) {
      lastError = error;
    }
    await sleep(100);
  }

  const suffix = lastError ? ` Last error: ${lastError.message}` : '';
  throw new Error(`Timed out waiting for ${label}.${suffix}`);
}

async function forceGarbageCollection(cdp, sessionId, passes) {
  await cdp.send('Runtime.discardConsoleEntries', {}, sessionId).catch(() => {});
  for (let index = 0; index < passes; index += 1) {
    await cdp.send('HeapProfiler.collectGarbage', {}, sessionId).catch(() => {});
    await sleep(50);
  }
}

async function findOrCreateTarget(cdp, url, createTargetIfMissing) {
  const urlPrefix = normalizeUrlPrefix(url);
  const { targetInfos } = await cdp.send('Target.getTargets');
  const candidates = targetInfos.filter((target) => {
    if (target.type !== 'page') return false;
    if (!target.url) return false;
    return normalizeUrlPrefix(target.url).startsWith(urlPrefix);
  });

  if (candidates[0]) return { ...candidates[0], createdByScript: false };
  if (!createTargetIfMissing) {
    throw new Error(`No existing Chrome tab matches ${url}`);
  }

  const created = await cdp.send('Target.createTarget', { url });
  return { targetId: created.targetId, type: 'page', url, createdByScript: true };
}

async function attachToTarget(cdp, targetId) {
  const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: true });
  await cdp.send('Runtime.enable', {}, sessionId);
  await cdp.send('Page.enable', {}, sessionId);
  await cdp.send('DOM.enable', {}, sessionId);
  await cdp.send('Performance.enable', {}, sessionId);
  await cdp.send('HeapProfiler.enable', {}, sessionId).catch(() => {});
  await cdp.send('Memory.enable', {}, sessionId).catch(() => {});
  await cdp.send('Network.enable', {}, sessionId).catch(() => {});
  return sessionId;
}

function createNetworkTracker() {
  const requests = [];
  const blocked = [];
  const suspiciousPatterns = [
    /\/doc\/[^/]+\/c(?:[/?#]|$)/i,
    /CommandService/i,
    /ConvertService/i,
    /coauthor/i,
    /websocket/i,
    /sockjs/i,
  ];

  return {
    attach(cdp, sessionId, appUrl) {
      const appOrigin = new URL(appUrl).origin;
      cdp.on('Network.requestWillBeSent', (params, eventSessionId) => {
        if (eventSessionId !== sessionId) return;
        const url = params.request?.url ?? '';
        const type = params.type ?? 'Other';
        const sameOrigin = url.startsWith(`${appOrigin}/`) || url === appOrigin;
        const localScheme = /^(blob:|data:|about:|chrome-extension:|devtools:)/.test(url);
        const suspicious = suspiciousPatterns.some((pattern) => pattern.test(url));
        if (suspicious || (!sameOrigin && !localScheme)) {
          requests.push({
            timestamp: new Date().toISOString(),
            method: params.request?.method,
            type,
            url,
            suspicious,
            sameOrigin,
          });
        }
      });
      cdp.on('Network.loadingFailed', (params, eventSessionId) => {
        if (eventSessionId !== sessionId) return;
        if (params.blockedReason || params.errorText) {
          blocked.push({
            timestamp: new Date().toISOString(),
            requestId: params.requestId,
            type: params.type,
            blockedReason: params.blockedReason,
            errorText: params.errorText,
          });
        }
      });
    },
    summary() {
      return {
        requestCount: requests.length,
        suspiciousCount: requests.filter((request) => request.suspicious).length,
        requests: requests.slice(-50),
        blocked: blocked.slice(-50),
      };
    },
  };
}

async function setHardResetOnClose(cdp, sessionId, enabled) {
  if (!enabled) return;
  await evaluate(
    cdp,
    sessionId,
    `(() => {
      const toggle = document.querySelector('#hard-reset-toggle');
      if (toggle) toggle.checked = true;
      return Boolean(toggle);
    })()`,
  );
}

async function setOpenMode(cdp, sessionId, mode) {
  const result = await evaluate(
    cdp,
    sessionId,
    `(() => {
      const mode = ${JSON.stringify(mode)};
      if (window.__officeDemo?.setMode) {
        window.__officeDemo.setMode(mode);
        return { ok: true, via: 'demo-api' };
      }
      const input = document.querySelector('input[name="open-mode"][value="' + mode + '"]');
      if (!input) return { ok: false, reason: 'missing mode input' };
      input.checked = true;
      input.dispatchEvent(new Event('change', { bubbles: true }));
      return { ok: true, via: 'radio' };
    })()`,
    { userGesture: true },
  );

  if (!result?.ok) {
    throw new Error(`Failed to set open mode ${mode}: ${result?.reason ?? 'unknown reason'}`);
  }
}

function groupWorkloadItemsByMode(items) {
  const groups = new Map();
  for (const item of items) {
    const mode = item.mode ?? 'edit';
    if (!groups.has(mode)) groups.set(mode, []);
    groups.get(mode).push(item);
  }
  return groups;
}

async function waitForOfficeEditorsReady(cdp, sessionId, expectedRecordCount, timeoutMs) {
  const started = Date.now();
  await waitForExpression(
    cdp,
    sessionId,
    `(() => {
      const records = Array.from(window.__officeDemo?.records ?? []);
      if (records.length < ${expectedRecordCount}) return false;
      return records.every((record) => record.instance?.getState?.().status === 'ready');
    })()`,
    timeoutMs,
    `${expectedRecordCount} office editor instance(s) to reach document ready`,
  );
  return Date.now() - started;
}

function emptyDemoExpression(previousBootId, requireBootChange) {
  const previousBootIdLiteral = JSON.stringify(previousBootId);
  return `(() => {
    const button = document.querySelector('#new-word-button');
    const rect = button?.getBoundingClientRect();
    const bootId = window.__officeDemo?.bootId ?? null;
    const bootReady = ${requireBootChange ? `Boolean(bootId && bootId !== ${previousBootIdLiteral})` : 'true'};
    return bootReady &&
      (window.__officeDemo?.records?.length ?? 0) === 0 &&
      document.querySelectorAll('iframe[name="frameEditor"]').length === 0 &&
      Boolean(button && rect && rect.width > 0 && rect.height > 0);
  })()`;
}

async function waitForEmptyDemo(cdp, sessionId, timeoutMs, previousBootId = null, requireBootChange = false) {
  await waitForExpression(
    cdp,
    sessionId,
    emptyDemoExpression(previousBootId, requireBootChange),
    timeoutMs,
    requireBootChange ? 'hard reset to finish and empty demo to return' : 'empty demo after editor close',
  );

  if (!requireBootChange) return;

  await sleep(RESET_STABILITY_MS);
  await waitForExpression(
    cdp,
    sessionId,
    emptyDemoExpression(previousBootId, true),
    timeoutMs,
    'stable empty demo after hard reset',
  );
}

async function openOfficeDocuments(cdp, sessionId, items, timeoutMs) {
  const beforeFrameCount = await evaluate(
    cdp,
    sessionId,
    `document.querySelectorAll('iframe[name="frameEditor"]').length`,
  );
  let expectedFrameCount = beforeFrameCount;

  for (const item of items) {
    await setOpenMode(cdp, sessionId, item.mode ?? 'edit');
    const format = item.format;
    const selector = FORMAT_SELECTORS[format];
    await waitForExpression(
      cdp,
      sessionId,
      `(() => {
        const button = document.querySelector(${JSON.stringify(selector)});
        if (!button) return false;
        const rect = button.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      })()`,
      timeoutMs,
      `${selector} to become visible`,
    );

    const clickResult = await evaluate(
      cdp,
      sessionId,
      `(() => {
        const button = document.querySelector(${JSON.stringify(selector)});
        if (!button) return { ok: false, reason: 'missing button' };
        button.click();
        return { ok: true };
      })()`,
      { userGesture: true },
    );

    if (!clickResult?.ok) {
      throw new Error(`Failed to click ${selector}: ${clickResult?.reason ?? 'unknown reason'}`);
    }

    expectedFrameCount += 1;
    await waitForExpression(
      cdp,
      sessionId,
      `document.querySelectorAll('iframe[name="frameEditor"]').length >= ${expectedFrameCount}`,
      timeoutMs,
      `${expectedFrameCount - beforeFrameCount} editor iframe(s)`,
    );
  }
}

async function openLocalFileGroup(cdp, sessionId, filePaths, expectedFrameCount, timeoutMs) {
  await waitForExpression(
    cdp,
    sessionId,
    `(() => {
      const button = document.querySelector('#upload-button');
      const input = document.querySelector('input[type="file"]');
      const rect = button?.getBoundingClientRect();
      return Boolean(button && input && rect && rect.width > 0 && rect.height > 0);
    })()`,
    timeoutMs,
    'upload button and file input',
  );

  const clickResult = await evaluate(
    cdp,
    sessionId,
    `(() => {
      const input = document.querySelector('input[type="file"]');
      const button = document.querySelector('#upload-button');
      if (!input || !button) return { ok: false, reason: 'missing upload controls' };
      const originalClick = HTMLInputElement.prototype.click;
      HTMLInputElement.prototype.click = function (...args) {
        if (this === input) return undefined;
        return originalClick.apply(this, args);
      };
      try {
        button.click();
        return { ok: true };
      } finally {
        HTMLInputElement.prototype.click = originalClick;
      }
    })()`,
    { userGesture: true },
  );

  if (!clickResult?.ok) {
    throw new Error(`Failed to prepare upload input: ${clickResult?.reason ?? 'unknown reason'}`);
  }

  const { root } = await cdp.send('DOM.getDocument', { depth: -1, pierce: true }, sessionId);
  const { nodeId } = await cdp.send('DOM.querySelector', { nodeId: root.nodeId, selector: 'input[type="file"]' }, sessionId);
  if (!nodeId) {
    throw new Error('Failed to find file input node');
  }

  await cdp.send('DOM.setFileInputFiles', { nodeId, files: filePaths }, sessionId);
  await evaluate(
    cdp,
    sessionId,
    `(() => {
      const input = document.querySelector('input[type="file"]');
      if (!input) return false;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    })()`,
    { userGesture: true },
  );

  await waitForExpression(
    cdp,
    sessionId,
    `document.querySelectorAll('iframe[name="frameEditor"]').length >= ${expectedFrameCount}`,
    timeoutMs,
    `${filePaths.map((filePath) => basename(filePath)).join(', ')} editor iframe(s)`,
  );
}

async function openLocalFiles(cdp, sessionId, items, timeoutMs) {
  const beforeFrameCount = await evaluate(
    cdp,
    sessionId,
    `document.querySelectorAll('iframe[name="frameEditor"]').length`,
  );
  let expectedFrameCount = beforeFrameCount;

  for (const [mode, modeItems] of groupWorkloadItemsByMode(items)) {
    await setOpenMode(cdp, sessionId, mode);
    const filePaths = modeItems.map((item) => item.filePath);
    expectedFrameCount += filePaths.length;
    await openLocalFileGroup(cdp, sessionId, filePaths, expectedFrameCount, timeoutMs);
  }
}

async function openWorkload(cdp, sessionId, workload, timeoutMs) {
  const items = workload.items ?? [workload];
  const beforeRecordCount = await evaluate(cdp, sessionId, `window.__officeDemo?.records?.length ?? 0`);
  const fileItems = items.filter((item) => item.filePath);
  const formatItems = items.filter((item) => !item.filePath);

  if (fileItems.length) {
    await openLocalFiles(cdp, sessionId, fileItems, timeoutMs);
  }

  if (formatItems.length) {
    await openOfficeDocuments(cdp, sessionId, formatItems, timeoutMs);
  }

  const readyWaitMs = await waitForOfficeEditorsReady(cdp, sessionId, beforeRecordCount + items.length, timeoutMs);
  workload.readyWaitMs = readyWaitMs;
  return { readyWaitMs };
}

async function closeOfficeDocument(cdp, sessionId, timeoutMs, closeMode, waitForHardReset) {
  const beforeFrameCount = await evaluate(
    cdp,
    sessionId,
    `document.querySelectorAll('iframe[name="frameEditor"]').length`,
  );
  const previousBootId = waitForHardReset
    ? await evaluate(cdp, sessionId, `window.__officeDemo?.bootId ?? null`)
    : null;
  if (closeMode === 'direct') {
    await evaluate(
      cdp,
      sessionId,
      `(() => {
        window.__officeDemo?.closeAll?.();
        return true;
      })()`,
      { userGesture: true },
    );
  } else if (closeMode === 'app') {
    const closeResult = await evaluate(
      cdp,
      sessionId,
      `(() => {
        const buttons = Array.from(document.querySelectorAll('[data-action="close"]'));
        const button = buttons.at(-1);
        if (!button) return { ok: false, reason: 'missing app close button' };
        button.click();
        return { ok: true };
      })()`,
      { userGesture: true },
    );

    if (!closeResult?.ok) {
      throw new Error(`Failed to close through app control: ${closeResult?.reason ?? 'unknown reason'}`);
    }
  } else {
    await evaluate(
      cdp,
      sessionId,
      `(() => {
        const buttons = Array.from(document.querySelectorAll('[data-action="close"]'));
        buttons.at(-1)?.click();
        return true;
      })()`,
      { userGesture: true },
    );
  }

  if (waitForHardReset) {
    await waitForEmptyDemo(cdp, sessionId, timeoutMs, previousBootId, true);
  } else {
    await waitForExpression(
      cdp,
      sessionId,
      `(() => {
        const button = document.querySelector('#new-word-button');
        const rect = button?.getBoundingClientRect();
        return document.querySelectorAll('iframe[name="frameEditor"]').length < ${beforeFrameCount} &&
          Boolean(button && rect && rect.width > 0 && rect.height > 0);
      })()`,
      timeoutMs,
      'home screen after editor close',
    );
  }
}

async function closeOfficeDocuments(cdp, sessionId, timeoutMs, closeMode, waitForHardReset) {
  const beforeFrameCount = await evaluate(
    cdp,
    sessionId,
    `document.querySelectorAll('iframe[name="frameEditor"]').length`,
  );
  if (beforeFrameCount === 0) return;
  const previousBootId = waitForHardReset
    ? await evaluate(cdp, sessionId, `window.__officeDemo?.bootId ?? null`)
    : null;

  if (closeMode === 'direct') {
    await evaluate(
      cdp,
      sessionId,
      `(() => {
        window.__officeDemo?.closeAll?.();
        return true;
      })()`,
      { userGesture: true },
    );
  } else if (closeMode === 'app') {
    const closeResult = await evaluate(
      cdp,
      sessionId,
      `(() => {
        const button = document.querySelector('#close-all-button');
        if (!button) return { ok: false, reason: 'missing close all button' };
        button.click();
        return { ok: true };
      })()`,
      { userGesture: true },
    );

    if (!closeResult?.ok) {
      throw new Error(`Failed to close all through app control: ${closeResult?.reason ?? 'unknown reason'}`);
    }
  } else {
    await evaluate(
      cdp,
      sessionId,
      `(() => {
        const buttons = Array.from(document.querySelectorAll('[data-action="close"]')).reverse();
        for (const button of buttons) button.click();
        return true;
      })()`,
      { userGesture: true },
    );
  }

  await waitForEmptyDemo(cdp, sessionId, timeoutMs, previousBootId, waitForHardReset);
}

async function measurePage(cdp, sessionId, browserRootPid, iteration, phase, workload, gcPasses) {
  if (gcPasses > 0) {
    await forceGarbageCollection(cdp, sessionId, gcPasses);
  }

  const [{ metrics }, domCounters, processInfo, pageState] = await Promise.all([
    cdp.send('Performance.getMetrics', {}, sessionId),
    cdp.send('Memory.getDOMCounters', {}, sessionId).catch(() => null),
    cdp.send('SystemInfo.getProcessInfo').catch(() => null),
    evaluate(
      cdp,
      sessionId,
      `(() => ({
        href: window.location.href,
        hash: window.location.hash,
        hasEditor: (window.__officeDemo?.records?.length ?? 0) > 0,
        activeOfficeEditors: window.__officeDemo?.records?.length ?? 0,
        activeOfficeEditorModes: Array.from(window.__officeDemo?.records ?? []).map((record) => record.instance?.getState?.().mode),
        iframeCount: document.querySelectorAll('iframe').length,
        frameEditorCount: document.querySelectorAll('iframe[name="frameEditor"]').length,
        elementCount: document.querySelectorAll('*').length,
        activeElement: document.activeElement?.tagName ?? null,
        appChildren: Array.from(document.querySelector('#app')?.children ?? []).map((child) => child.tagName),
      }))()`,
    ).catch(() => null),
  ]);

  const rss = collectBrowserRssMb(browserRootPid);
  const jsHeapUsedBytes = metricValue(metrics, 'JSHeapUsedSize');
  const jsHeapTotalBytes = metricValue(metrics, 'JSHeapTotalSize');

  return {
    iteration,
    phase,
    format: getWorkloadFormat(workload),
    mode: getWorkloadMode(workload),
    readyWaitMs: getWorkloadReadyWaitMs(workload),
    fileName: getWorkloadFileName(workload),
    filePath: getWorkloadFilePath(workload),
    timestamp: new Date().toISOString(),
    jsHeapUsedMb: bytesToMb(jsHeapUsedBytes),
    jsHeapTotalMb: bytesToMb(jsHeapTotalBytes),
    documents: domCounters?.documents ?? metricValue(metrics, 'Documents'),
    nodes: domCounters?.nodes ?? metricValue(metrics, 'Nodes'),
    jsEventListeners: domCounters?.jsEventListeners ?? metricValue(metrics, 'JSEventListeners'),
    layoutObjects: metricValue(metrics, 'LayoutObjects'),
    browserRssMb: rss.totalMb,
    browserProcessCount: rss.processCount,
    browserRssBreakdown: rss.breakdown,
    browserTopProcesses: rss.topProcesses,
    cdpProcessInfo: processInfo?.processInfo ?? null,
    pageState,
  };
}

function breakdownRssMb(sample, type) {
  return sample.browserRssBreakdown?.[type]?.rssMb ?? null;
}

function logSample(sample, baseline) {
  const heapDelta =
    baseline && Number.isFinite(sample.jsHeapUsedMb) && Number.isFinite(baseline.jsHeapUsedMb)
      ? sample.jsHeapUsedMb - baseline.jsHeapUsedMb
      : null;
  const rssDelta =
    baseline && Number.isFinite(sample.browserRssMb) && Number.isFinite(baseline.browserRssMb)
      ? sample.browserRssMb - baseline.browserRssMb
      : null;
  const readyLabel = Number.isFinite(sample.readyWaitMs) ? `${sample.readyWaitMs}ms` : 'n/a';

  console.log(
    [
      `${String(sample.iteration).padStart(3, '0')} ${sample.phase.padEnd(9)}`,
      `${sample.format || 'base'}`.padEnd(4),
      `mode=${sample.mode || 'base'}`,
      `ready=${readyLabel}`,
      `heap=${formatMb(sample.jsHeapUsedMb)} MB`,
      `heapDelta=${formatMb(heapDelta)} MB`,
      `docs=${sample.documents ?? 'n/a'}`,
      `nodes=${sample.nodes ?? 'n/a'}`,
      `listeners=${sample.jsEventListeners ?? 'n/a'}`,
      `rss=${formatMb(sample.browserRssMb)} MB`,
      `rssDelta=${formatMb(rssDelta)} MB`,
      `renderer=${formatMb(breakdownRssMb(sample, 'renderer'))} MB`,
      `gpu=${formatMb(breakdownRssMb(sample, 'gpu'))} MB`,
      `utility=${formatMb(breakdownRssMb(sample, 'utility'))} MB`,
    ].join(' | '),
  );
}

function toCsv(samples) {
  const columns = [
    'iteration',
    'phase',
    'format',
    'mode',
    'readyWaitMs',
    'fileName',
    'timestamp',
    'jsHeapUsedMb',
    'jsHeapTotalMb',
    'documents',
    'nodes',
    'jsEventListeners',
    'layoutObjects',
    'browserRssMb',
    'browserProcessCount',
    'browserRssBrowserMb',
    'browserRssRendererMb',
    'browserRssGpuMb',
    'browserRssUtilityMb',
    'browserRssNetworkMb',
    'browserRssOtherMb',
    'pageHasEditor',
    'pageActiveModes',
    'pageIframeCount',
    'pageFrameEditorCount',
    'pageElementCount',
    'url',
  ];

  const escape = (value) => {
    if (value === null || value === undefined) return '';
    const text = String(value);
    if (!/[",\n]/.test(text)) return text;
    return `"${text.replaceAll('"', '""')}"`;
  };

  const rowValue = (sample, column) => {
    switch (column) {
      case 'browserRssBrowserMb':
        return breakdownRssMb(sample, 'browser');
      case 'browserRssRendererMb':
        return breakdownRssMb(sample, 'renderer');
      case 'browserRssGpuMb':
        return breakdownRssMb(sample, 'gpu');
      case 'browserRssUtilityMb':
        return breakdownRssMb(sample, 'utility');
      case 'browserRssNetworkMb':
        return breakdownRssMb(sample, 'network');
      case 'browserRssOtherMb':
        return breakdownRssMb(sample, 'other');
      case 'pageHasEditor':
        return sample.pageState?.hasEditor;
      case 'pageActiveModes':
        return sample.pageState?.activeOfficeEditorModes?.join(' + ');
      case 'pageIframeCount':
        return sample.pageState?.iframeCount;
      case 'pageFrameEditorCount':
        return sample.pageState?.frameEditorCount;
      case 'pageElementCount':
        return sample.pageState?.elementCount;
      case 'url':
        return sample.pageState?.href;
      default:
        return sample[column];
    }
  };

  return [
    columns.join(','),
    ...samples.map((sample) => columns.map((column) => escape(rowValue(sample, column))).join(',')),
  ].join('\n');
}

function maxFinite(values) {
  const finite = values.filter(Number.isFinite);
  return finite.length ? Math.max(...finite) : null;
}

function analyze(samples, networkSummary) {
  const baseline = samples.find((sample) => sample.phase === 'baseline');
  const closed = samples.filter((sample) => sample.phase === 'closed');
  const opened = samples.filter((sample) => sample.phase === 'opened');
  const settled = samples.find((sample) => sample.phase === 'settled');
  const final = samples.find((sample) => sample.phase === 'final');
  const firstClosed = closed[0];
  const lastClosed = closed.at(-1);
  const finalClosed = settled ?? final ?? lastClosed;

  if (!baseline || !firstClosed || !finalClosed) {
    return { verdict: 'Not enough samples to analyze.', baseline, firstClosed, finalClosed };
  }

  const denominator = Math.max(closed.length - 1, 1);
  const closedHeapDeltaMb = finalClosed.jsHeapUsedMb - baseline.jsHeapUsedMb;
  const closedHeapDeltaSinceFirstMb = finalClosed.jsHeapUsedMb - firstClosed.jsHeapUsedMb;
  const closedRssDeltaMb =
    Number.isFinite(finalClosed.browserRssMb) && Number.isFinite(baseline.browserRssMb)
      ? finalClosed.browserRssMb - baseline.browserRssMb
      : null;
  const closedRssDeltaSinceFirstMb =
    Number.isFinite(finalClosed.browserRssMb) && Number.isFinite(firstClosed.browserRssMb)
      ? finalClosed.browserRssMb - firstClosed.browserRssMb
      : null;
  const heapGrowthPerCycleMb = closedHeapDeltaSinceFirstMb / denominator;
  const rssGrowthPerCycleMb = Number.isFinite(closedRssDeltaSinceFirstMb)
    ? closedRssDeltaSinceFirstMb / denominator
    : null;
  const peakClosedRssMb = maxFinite(closed.map((sample) => sample.browserRssMb));
  const peakOpenRssMb = maxFinite(opened.map((sample) => sample.browserRssMb));
  const peakOpenHeapMb = maxFinite(opened.map((sample) => sample.jsHeapUsedMb));
  const finalNodeDelta = finalClosed.nodes - baseline.nodes;
  const finalListenerDelta = finalClosed.jsEventListeners - baseline.jsEventListeners;
  const finalDocumentDelta = finalClosed.documents - baseline.documents;
  const finalPageClean =
    finalClosed.pageState &&
    !finalClosed.pageState.hasEditor &&
    finalClosed.pageState.frameEditorCount === 0 &&
    finalClosed.pageState.iframeCount === 0 &&
    finalClosed.pageState.elementCount <= 100;

  let verdict = 'closed-state heap/DOM counters plateau after forced GC';
  const suspiciousHeapGrowth = closedHeapDeltaMb > 30 && heapGrowthPerCycleMb > 3;
  if (
    (!finalPageClean && (finalDocumentDelta > 4 || finalNodeDelta > 2_000)) ||
    finalNodeDelta > 5_000 ||
    finalListenerDelta > 500 ||
    suspiciousHeapGrowth
  ) {
    verdict = 'possible retained page objects after close; inspect final closed DOM counters';
  }
  if (
    Number.isFinite(rssGrowthPerCycleMb) &&
    Number.isFinite(closedRssDeltaMb) &&
    closedRssDeltaMb > 120 &&
    (rssGrowthPerCycleMb > 8 || closedRssDeltaSinceFirstMb > 120)
  ) {
    verdict = 'possible retained non-JS/native memory; Chrome RSS keeps rising after close';
  }
  if (networkSummary.suspiciousCount > 0) {
    verdict = `${verdict}; suspicious DocumentServer-like requests were observed`;
  }

  return {
    verdict,
    baselineHeapMb: baseline.jsHeapUsedMb,
    firstClosedHeapMb: firstClosed.jsHeapUsedMb,
    finalClosedHeapMb: finalClosed.jsHeapUsedMb,
    closedHeapDeltaMb,
    closedHeapDeltaSinceFirstMb,
    heapGrowthPerCycleMb,
    baselineRssMb: baseline.browserRssMb,
    firstClosedRssMb: firstClosed.browserRssMb,
    finalClosedRssMb: finalClosed.browserRssMb,
    peakClosedRssMb,
    peakOpenRssMb,
    peakOpenHeapMb,
    closedRssDeltaMb,
    closedRssDeltaSinceFirstMb,
    rssGrowthPerCycleMb,
    finalDocumentDelta,
    finalNodeDelta,
    finalListenerDelta,
    finalRssBreakdown: finalClosed.browserRssBreakdown,
    finalTopProcesses: finalClosed.browserTopProcesses,
    networkSummary,
  };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const cdp = new CdpClient(options.browserWs);
  await cdp.connect();

  const version = await cdp.send('Browser.getVersion').catch(() => null);
  const target = await findOrCreateTarget(cdp, options.url, options.createTargetIfMissing);
  const sessionId = await attachToTarget(cdp, target.targetId);
  const networkTracker = createNetworkTracker();
  networkTracker.attach(cdp, sessionId, options.url);

  const debugPort = inferDebugPort(options.browserWs);
  const browserRootPid = options.chromePid ?? findChromeRootPid(debugPort);
  const random = createRandom(options.seed);
  const workloads = options.filePaths.length
    ? options.filePaths.map((filePath) => ({ filePath, format: getFileFormat(filePath) }))
    : options.formats.map((format) => ({ format }));
  const samples = [];

  try {
    await cdp.send('Page.navigate', { url: options.url }, sessionId);
    await waitForExpression(
      cdp,
      sessionId,
      `Boolean(document.querySelector('#new-word-button'))`,
      options.timeoutMs,
      'app home screen',
    );
    await sleep(options.homeDwellMs);

    console.log(`Connected to ${version?.product ?? options.browserWs}`);
    console.log(`Target: ${target.targetId}`);
    console.log(`Browser PID for RSS: ${browserRootPid ?? 'n/a'}`);
    console.log(
      `Running ${options.iterations} random ${workloads.map(getWorkloadLabel).join(' / ')} cycles, modes=${options.modes.join(',')}, batch=${options.batchSize}, stay=${options.stayMs}ms, home=${options.homeDwellMs}ms`,
    );
    console.log(`Seed: ${options.seed}`);

    const baseline = await measurePage(cdp, sessionId, browserRootPid, 0, 'baseline', null, options.gcPasses);
    samples.push(baseline);
    logSample(baseline, baseline);

    for (let iteration = 1; iteration <= options.iterations; iteration += 1) {
      const workload = pickWorkloadBatch(random, workloads, options.batchSize, options.modes);
      console.log(`\nCycle ${iteration}/${options.iterations}: ${getWorkloadLabel(workload)}`);

      await setHardResetOnClose(cdp, sessionId, options.hardResetOnClose);
      const { readyWaitMs } = await openWorkload(cdp, sessionId, workload, options.timeoutMs);
      console.log(`  ready after ${readyWaitMs}ms`);
      await sleep(options.stayMs);

      if (options.openedSampleInterval > 0 && iteration % options.openedSampleInterval === 0) {
        const opened = await measurePage(cdp, sessionId, browserRootPid, iteration, 'opened', workload, 0);
        samples.push(opened);
        logSample(opened, baseline);
      }

      if (workload.items) {
        await closeOfficeDocuments(cdp, sessionId, options.timeoutMs, options.closeMode, options.hardResetOnClose);
      } else {
        await closeOfficeDocument(cdp, sessionId, options.timeoutMs, options.closeMode, options.hardResetOnClose);
      }
      await sleep(options.homeDwellMs);

      const closed = await measurePage(cdp, sessionId, browserRootPid, iteration, 'closed', workload, options.gcPasses);
      samples.push(closed);
      logSample(closed, baseline);
    }

    const finalSample = await measurePage(
      cdp,
      sessionId,
      browserRootPid,
      options.iterations,
      'final',
      null,
      options.finalGcPasses,
    );
    samples.push(finalSample);
    logSample(finalSample, baseline);

    if (options.finalSettleMs > 0) {
      await sleep(options.finalSettleMs);
      const settledSample = await measurePage(
        cdp,
        sessionId,
        browserRootPid,
        options.iterations,
        'settled',
        null,
        options.finalGcPasses,
      );
      samples.push(settledSample);
      logSample(settledSample, baseline);
    }

    const networkSummary = networkTracker.summary();
    const analysis = analyze(samples, networkSummary);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const outputDir = resolve(ROOT, 'test-results', 'memory');
    mkdirSync(outputDir, { recursive: true });
    const jsonPath = resolve(outputDir, `chrome-cdp-memory-${timestamp}.json`);
    const csvPath = resolve(outputDir, `chrome-cdp-memory-${timestamp}.csv`);

    writeFileSync(jsonPath, JSON.stringify({ options, browserVersion: version, analysis, samples }, null, 2));
    writeFileSync(csvPath, `${toCsv(samples)}\n`);

    console.log('\nAnalysis');
    console.log(`  verdict: ${analysis.verdict}`);
    console.log(`  final closed heap delta vs blank: ${formatMb(analysis.closedHeapDeltaMb)} MB`);
    console.log(`  closed heap delta after first close: ${formatMb(analysis.closedHeapDeltaSinceFirstMb)} MB`);
    console.log(`  closed heap growth/cycle after first close: ${formatMb(analysis.heapGrowthPerCycleMb)} MB`);
    console.log(`  final closed RSS delta vs blank: ${formatMb(analysis.closedRssDeltaMb)} MB`);
    console.log(`  closed RSS delta after first close: ${formatMb(analysis.closedRssDeltaSinceFirstMb)} MB`);
    console.log(`  closed RSS growth/cycle after first close: ${formatMb(analysis.rssGrowthPerCycleMb)} MB`);
    console.log(`  final document delta: ${analysis.finalDocumentDelta ?? 'n/a'}`);
    console.log(`  final DOM node delta: ${analysis.finalNodeDelta ?? 'n/a'}`);
    console.log(`  final listener delta: ${analysis.finalListenerDelta ?? 'n/a'}`);
    console.log(`  suspicious network requests: ${networkSummary.suspiciousCount}/${networkSummary.requestCount}`);
    if (analysis.finalRssBreakdown) {
      console.log('  final RSS breakdown:');
      for (const [type, detail] of Object.entries(analysis.finalRssBreakdown)) {
        console.log(`    ${type}: ${formatMb(detail.rssMb)} MB across ${detail.count} process(es)`);
      }
    }
    console.log(`  JSON: ${jsonPath}`);
    console.log(`  CSV:  ${csvPath}`);
  } finally {
    await cdp.send('Target.detachFromTarget', { sessionId }).catch(() => {});
    if (target.createdByScript && !options.keepTarget) {
      await cdp.send('Target.closeTarget', { targetId: target.targetId }).catch(() => {});
    }
    cdp.close();
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack || error.message : error);
  process.exit(1);
});
