const MIB = 1024 * 1024;

export const CANONICAL_BROKER_ACCEPTANCE_DEFAULTS = Object.freeze({
  requiredLifecycleIterations: 100,
  requiredConcurrentOrigins: 3,
  maxBrokerReservedBytes: 64 * MIB,
  maxRecoveryMs: 30_000,
  maxFinalHeapDeltaMb: 32,
  maxFinalRssDeltaMb: 128,
  stableWindowCycles: 20,
  maxStableHeapRangeMb: 16,
  maxStableRssRangeMb: 96,
  maxStableHeapSlopeMbPerCycle: 0.25,
  maxStableRssSlopeMbPerCycle: 1,
});

function finite(values) {
  return values.filter(Number.isFinite);
}

export function percentile(values, percentileValue) {
  const sorted = finite(values).sort((left, right) => left - right);
  if (!sorted.length) return null;
  if (!Number.isFinite(percentileValue) || percentileValue < 0 || percentileValue > 1) {
    throw new TypeError('percentile must be between zero and one');
  }
  const index = Math.max(0, Math.ceil(sorted.length * percentileValue) - 1);
  return sorted[index];
}

export function linearSlope(values) {
  const points = values.map((value, index) => ({ x: index, y: value })).filter((point) => Number.isFinite(point.y));
  if (points.length < 2) return null;
  const xMean = points.reduce((sum, point) => sum + point.x, 0) / points.length;
  const yMean = points.reduce((sum, point) => sum + point.y, 0) / points.length;
  const denominator = points.reduce((sum, point) => sum + (point.x - xMean) ** 2, 0);
  if (denominator === 0) return 0;
  return points.reduce((sum, point) => sum + (point.x - xMean) * (point.y - yMean), 0) / denominator;
}

function range(values) {
  const valuesToMeasure = finite(values);
  return valuesToMeasure.length ? Math.max(...valuesToMeasure) - Math.min(...valuesToMeasure) : null;
}

function check(id, actual, expected, pass, detail = null) {
  return { id, actual, expected, pass: pass === true, detail };
}

export function analyzeLifecycleReport(report, overrides = {}) {
  const thresholds = { ...CANONICAL_BROKER_ACCEPTANCE_DEFAULTS, ...overrides };
  const closed = (report?.samples || []).filter((sample) => sample.phase === 'closed');
  const stable = closed.slice(-thresholds.stableWindowCycles);
  const heapValues = stable.map((sample) => sample.jsHeapUsedMb);
  const rssValues = stable.map((sample) => sample.browserRssMb);
  const analysis = report?.analysis || {};
  const checks = [
    check(
      'lifecycle-iterations',
      closed.length,
      `>= ${thresholds.requiredLifecycleIterations}`,
      closed.length >= thresholds.requiredLifecycleIterations,
    ),
    check(
      'final-heap-delta',
      analysis.closedHeapDeltaMb,
      `<= ${thresholds.maxFinalHeapDeltaMb} MiB`,
      Number.isFinite(analysis.closedHeapDeltaMb) && analysis.closedHeapDeltaMb <= thresholds.maxFinalHeapDeltaMb,
    ),
    check(
      'final-rss-delta',
      analysis.closedRssDeltaMb,
      `<= ${thresholds.maxFinalRssDeltaMb} MiB`,
      Number.isFinite(analysis.closedRssDeltaMb) && analysis.closedRssDeltaMb <= thresholds.maxFinalRssDeltaMb,
    ),
    check(
      'stable-heap-range',
      range(heapValues),
      `<= ${thresholds.maxStableHeapRangeMb} MiB over final ${thresholds.stableWindowCycles} cycles`,
      stable.length >= thresholds.stableWindowCycles &&
        Number.isFinite(range(heapValues)) &&
        range(heapValues) <= thresholds.maxStableHeapRangeMb,
    ),
    check(
      'stable-rss-range',
      range(rssValues),
      `<= ${thresholds.maxStableRssRangeMb} MiB over final ${thresholds.stableWindowCycles} cycles`,
      stable.length >= thresholds.stableWindowCycles &&
        Number.isFinite(range(rssValues)) &&
        range(rssValues) <= thresholds.maxStableRssRangeMb,
    ),
    check(
      'stable-heap-slope',
      linearSlope(heapValues),
      `<= ${thresholds.maxStableHeapSlopeMbPerCycle} MiB/cycle`,
      stable.length >= thresholds.stableWindowCycles &&
        Number.isFinite(linearSlope(heapValues)) &&
        linearSlope(heapValues) <= thresholds.maxStableHeapSlopeMbPerCycle,
    ),
    check(
      'stable-rss-slope',
      linearSlope(rssValues),
      `<= ${thresholds.maxStableRssSlopeMbPerCycle} MiB/cycle`,
      stable.length >= thresholds.stableWindowCycles &&
        Number.isFinite(linearSlope(rssValues)) &&
        linearSlope(rssValues) <= thresholds.maxStableRssSlopeMbPerCycle,
    ),
    check(
      'message-port-metrics-present',
      analysis.messagePortMetricsAvailable,
      true,
      analysis.messagePortMetricsAvailable === true,
    ),
    check(
      'message-ports-reclaimed',
      analysis.finalPortState ?? null,
      'all active ports, instances, and origin leases are zero',
      analysis.finalMessagePortsClean === true,
    ),
    check(
      'no-parent-refresh-loop',
      analysis.unexpectedPageRefreshes?.length ?? null,
      0,
      Array.isArray(analysis.unexpectedPageRefreshes) && analysis.unexpectedPageRefreshes.length === 0,
    ),
  ];
  return { pass: checks.every((item) => item.pass), thresholds, checks };
}

export function analyzeStartupPerformance(report, baselineP95Ms) {
  const readySamples = (report?.samples || [])
    .filter((sample) => sample.phase === 'closed')
    .map((sample) => sample.readyWaitMs);
  const readyP95Ms = percentile(readySamples, 0.95);
  const allowedExtraMs = Number.isFinite(baselineP95Ms) ? Math.max(300, baselineP95Ms * 0.1) : null;
  const extraP95Ms =
    Number.isFinite(readyP95Ms) && Number.isFinite(baselineP95Ms) ? Math.max(0, readyP95Ms - baselineP95Ms) : null;
  const checks = [
    check(
      'startup-baseline-present',
      baselineP95Ms,
      'a finite pre-recorded baseline P95',
      Number.isFinite(baselineP95Ms),
    ),
    check('startup-samples-present', readySamples.length, '> 0', readySamples.length > 0),
    check(
      'broker-startup-extra-p95',
      extraP95Ms,
      `<= max(300 ms, ${baselineP95Ms} ms × 10%) = ${allowedExtraMs} ms`,
      Number.isFinite(extraP95Ms) && Number.isFinite(allowedExtraMs) && extraP95Ms <= allowedExtraMs,
      { readyP95Ms, baselineP95Ms, allowedExtraMs },
    ),
  ];
  return { pass: checks.every((item) => item.pass), readyP95Ms, baselineP95Ms, allowedExtraMs, extraP95Ms, checks };
}

function metricSnapshotsByRole(snapshots, role) {
  return (snapshots || []).filter((snapshot) => snapshot?.role === role);
}

export function analyzeBrokerMetrics(input, overrides = {}) {
  const thresholds = { ...CANONICAL_BROKER_ACCEPTANCE_DEFAULTS, ...overrides };
  const canonicalSnapshots = metricSnapshotsByRole(input?.snapshots, 'canonical-service-worker');
  const editorSnapshots = metricSnapshotsByRole(input?.snapshots, 'editor-service-worker');
  const relaySnapshots = metricSnapshotsByRole(input?.snapshots, 'canonical-relay-frame');
  const canonicalMetrics = canonicalSnapshots.map((snapshot) => snapshot.canonical).filter(Boolean);
  const editorMetrics = editorSnapshots.map((snapshot) => snapshot.editor).filter(Boolean);
  const relayMetrics = relaySnapshots.map((snapshot) => snapshot.relay).filter(Boolean);
  // This is deliberately conservative: component-local peaks may have
  // occurred at different instants, but summing them cannot hide an aggregate
  // buffer-budget violation.
  const peakReservedBytes =
    Math.max(0, ...canonicalMetrics.map((metric) => metric.peakReservedBytes)) +
    editorMetrics.reduce((total, metric) => total + metric.peakReservedBytes, 0) +
    relayMetrics.reduce((total, metric) => total + metric.peakReservedBytes, 0);
  const finalCanonical = canonicalMetrics.at(-1) ?? null;
  const finalEditors = input?.finalEditorSnapshots || editorMetrics;
  const burst = input?.burstStatus;
  const checks = [
    check('canonical-metrics-present', canonicalMetrics.length, '>= 1', canonicalMetrics.length >= 1),
    check(
      'editor-metrics-present',
      editorMetrics.length,
      `>= ${thresholds.requiredConcurrentOrigins}`,
      editorMetrics.length >= thresholds.requiredConcurrentOrigins,
    ),
    check(
      'relay-metrics-present',
      relayMetrics.length,
      `>= ${thresholds.requiredConcurrentOrigins}`,
      relayMetrics.length >= thresholds.requiredConcurrentOrigins,
    ),
    check(
      'three-origin-ready',
      burst ?? null,
      `${thresholds.requiredConcurrentOrigins} ready editors on ${thresholds.requiredConcurrentOrigins} origins with no errors`,
      burst?.ready === thresholds.requiredConcurrentOrigins &&
        burst?.uniqueOrigins === thresholds.requiredConcurrentOrigins &&
        Array.isArray(burst?.errors) &&
        burst.errors.length === 0,
    ),
    check(
      'broker-buffer-high-water',
      peakReservedBytes,
      `<= ${thresholds.maxBrokerReservedBytes} bytes`,
      peakReservedBytes <= thresholds.maxBrokerReservedBytes,
    ),
    check(
      'canonical-reads-drained',
      finalCanonical,
      'activeReads=0 and reservedBytes=0',
      finalCanonical?.activeReads === 0 && finalCanonical?.reservedBytes === 0,
    ),
    check(
      'editor-reads-drained',
      finalEditors,
      'every editor is disconnected with activeRequests=0 and reservedBytes=0',
      finalEditors.length >= thresholds.requiredConcurrentOrigins &&
        finalEditors.every(
          (metric) =>
            metric.connectionStatus === 'disconnected' && metric.activeRequests === 0 && metric.reservedBytes === 0,
        ),
    ),
  ];
  return { pass: checks.every((item) => item.pass), thresholds, peakReservedBytes, checks };
}

export function analyzeRecoveryResults(results, overrides = {}) {
  const thresholds = { ...CANONICAL_BROKER_ACCEPTANCE_DEFAULTS, ...overrides };
  const requiredCases = new Set([
    'broker-iframe',
    'canonical-service-worker',
    'editor-service-worker',
    'message-port',
    'freeze-resume',
  ]);
  const providedCases = new Set((results || []).map((result) => result.name));
  const checks = [
    check(
      'all-recovery-cases-present',
      [...providedCases].sort(),
      [...requiredCases].sort(),
      [...requiredCases].every((name) => providedCases.has(name)),
    ),
    ...(results || []).map((result) =>
      check(
        `recovery-${result.name}`,
        result,
        `recovered or accurate-error within ${thresholds.maxRecoveryMs} ms`,
        (result.outcome === 'recovered' || result.outcome === 'accurate-error') &&
          Number.isFinite(result.elapsedMs) &&
          result.elapsedMs <= thresholds.maxRecoveryMs &&
          (result.outcome !== 'accurate-error' || typeof result.errorCode === 'string'),
      ),
    ),
  ];
  return { pass: checks.every((item) => item.pass), thresholds, checks };
}

function normalizedContentType(value) {
  return typeof value === 'string' ? value.split(';', 1)[0].trim().toLowerCase() : '';
}

function validDigest(value) {
  return typeof value === 'string' && /^[a-f0-9]{64}$/u.test(value);
}

function isEditorShellPath(path) {
  return (
    path === 'office-host.html' ||
    path === 'editor-shell-prime.html' ||
    (typeof path === 'string' && /^assets\/[a-zA-Z0-9._+-]+\.(?:css|js)$/u.test(path))
  );
}

export function analyzeReleaseIntegrity(input) {
  const assets = Array.isArray(input?.assets) ? input.assets : [];
  const expectedAssetCount = input?.expectedAssetCount;
  const checks = [
    check(
      'all-release-assets-present',
      assets.length,
      `exactly ${expectedAssetCount} manifest assets`,
      Number.isSafeInteger(expectedAssetCount) &&
        expectedAssetCount > 0 &&
        assets.length === expectedAssetCount &&
        new Set(assets.map((asset) => asset?.path)).size === expectedAssetCount,
    ),
    ...assets.map((asset) => {
      const expectedMime = normalizedContentType(asset?.expectedMime);
      const actualMime = normalizedContentType(asset?.contentType);
      const pass =
        typeof asset?.path === 'string' &&
        asset.path.length > 0 &&
        asset.status === 200 &&
        Number.isSafeInteger(asset.expectedBytes) &&
        asset.expectedBytes >= 0 &&
        asset.bytes === asset.expectedBytes &&
        asset.bytesSent === asset.expectedBytes &&
        (asset.expectedBytes === 0 ? asset.chunks === 0 : Number.isSafeInteger(asset.chunks) && asset.chunks > 0) &&
        asset.contentLength === asset.expectedBytes &&
        validDigest(asset.expectedSha256) &&
        asset.sha256 === asset.expectedSha256 &&
        expectedMime.length > 0 &&
        actualMime === expectedMime &&
        asset.acceptRanges === 'bytes' &&
        asset.terminalType === 'END';
      return check(
        `asset-${asset?.path ?? 'missing-path'}`,
        {
          status: asset?.status,
          bytes: asset?.bytes,
          bytesSent: asset?.bytesSent,
          chunks: asset?.chunks,
          sha256: asset?.sha256,
          contentType: actualMime,
          contentLength: asset?.contentLength,
          acceptRanges: asset?.acceptRanges,
          terminalType: asset?.terminalType,
        },
        {
          status: 200,
          bytes: asset?.expectedBytes,
          bytesSent: asset?.expectedBytes,
          chunks: asset?.expectedBytes === 0 ? 0 : '> 0',
          sha256: asset?.expectedSha256,
          contentType: expectedMime,
          contentLength: asset?.expectedBytes,
          acceptRanges: 'bytes',
          terminalType: 'END',
        },
        pass,
      );
    }),
  ];

  const full = input?.samples?.full;
  checks.push(
    check(
      'sample-full-200',
      full ?? null,
      '200 with exact manifest bytes, SHA-256, MIME, Content-Length, and END',
      full?.status === 200 &&
        typeof full?.path === 'string' &&
        !isEditorShellPath(full.path) &&
        full?.bytes === full?.expectedBytes &&
        full?.contentLength === full?.expectedBytes &&
        validDigest(full?.expectedSha256) &&
        full?.sha256 === full?.expectedSha256 &&
        normalizedContentType(full?.contentType) === normalizedContentType(full?.expectedMime) &&
        full?.acceptRanges === 'bytes' &&
        full?.terminalType === 'END',
    ),
  );

  const range = input?.samples?.range;
  checks.push(
    check(
      'sample-range-206',
      range ?? null,
      '206 with exact requested bytes, Content-Range, SHA-256, and identical direct-Broker bytes',
      range?.status === 206 &&
        typeof range?.path === 'string' &&
        !isEditorShellPath(range?.path) &&
        Number.isSafeInteger(range?.start) &&
        range.start > 0 &&
        Number.isSafeInteger(range?.end) &&
        range.end >= range.start &&
        range?.expectedSource === 'full-broker-stream-capture' &&
        Number.isSafeInteger(range?.bytes) &&
        range.bytes > 0 &&
        range.contentLength === range.bytes &&
        range.contentRange === range.expectedContentRange &&
        validDigest(range?.expectedSha256) &&
        range.sha256 === range.expectedSha256 &&
        range.directSha256 === range.expectedSha256 &&
        normalizedContentType(range?.contentType) === normalizedContentType(range?.expectedMime) &&
        range?.acceptRanges === 'bytes' &&
        range.terminalType === 'END',
    ),
  );

  const unsatisfied = input?.samples?.unsatisfied;
  checks.push(
    check(
      'sample-range-416',
      unsatisfied ?? null,
      '416 with zero body and bytes */<manifest length>',
      unsatisfied?.status === 416 &&
        unsatisfied?.bytes === 0 &&
        unsatisfied?.bytesSent === 0 &&
        unsatisfied?.contentLength === 0 &&
        unsatisfied?.contentRange === unsatisfied?.expectedContentRange &&
        unsatisfied?.acceptRanges === 'bytes' &&
        unsatisfied?.terminalType === 'END',
    ),
  );

  const cancellation = input?.samples?.cancellation;
  checks.push(
    check(
      'sample-cancellation-and-recovery',
      cancellation ?? null,
      'direct and editor AbortController cancellation drain within 30s, followed by exact one-byte recovery',
      cancellation?.status === 206 &&
        cancellation?.terminalType === 'CANCELLED' &&
        Number.isFinite(cancellation?.elapsedMs) &&
        cancellation.elapsedMs <= 30_000 &&
        Number.isSafeInteger(cancellation?.bytes) &&
        cancellation.bytes > 0 &&
        cancellation.bytesSent === cancellation.bytes &&
        Number.isSafeInteger(cancellation?.expectedBytes) &&
        cancellation.bytes < cancellation.expectedBytes &&
        cancellation?.editor?.status === 206 &&
        cancellation.editor.contentLength === cancellation.expectedBytes &&
        cancellation.editor.contentRange === cancellation.editor.expectedContentRange &&
        cancellation.editor.acceptRanges === 'bytes' &&
        Number.isSafeInteger(cancellation.editor.firstBytes) &&
        cancellation.editor.firstBytes > 0 &&
        cancellation.editor.firstBytes < cancellation.expectedBytes &&
        cancellation.editor.firstDone === false &&
        cancellation.editor.aborted === true &&
        cancellation.editor.errorName === 'AbortError' &&
        Number.isFinite(cancellation.editor.elapsedMs) &&
        cancellation.editor.elapsedMs <= 30_000 &&
        cancellation?.drain?.drained === true &&
        Number.isFinite(cancellation.drain.elapsedMs) &&
        cancellation.drain.elapsedMs <= 30_000 &&
        cancellation.drain.canonical?.activeReads === 0 &&
        cancellation.drain.canonical?.reservedBytes === 0 &&
        cancellation.drain.relay?.activeRequests === 0 &&
        cancellation.drain.relay?.reservedBytes === 0 &&
        cancellation.drain.editor?.activeRequests === 0 &&
        cancellation.drain.editor?.reservedBytes === 0 &&
        cancellation?.protocolRecovery?.status === 206 &&
        cancellation.protocolRecovery.bytes === 1 &&
        cancellation.protocolRecovery.bytesSent === 1 &&
        cancellation.protocolRecovery.contentLength === 1 &&
        cancellation.protocolRecovery.contentRange === cancellation.protocolRecovery.expectedContentRange &&
        cancellation.protocolRecovery.terminalType === 'END' &&
        cancellation?.recovery?.status === 206 &&
        cancellation.recovery.bytes === 1 &&
        cancellation.recovery.contentLength === 1 &&
        cancellation.recovery.contentRange === cancellation.recovery.expectedContentRange &&
        cancellation.recovery.acceptRanges === 'bytes' &&
        cancellation.recovery.expectedSource === 'full-broker-stream-capture' &&
        validDigest(cancellation.recovery.expectedSha256) &&
        cancellation.recovery.sha256 === cancellation.recovery.expectedSha256 &&
        cancellation.recovery.terminalType === 'END',
    ),
  );

  const wasm = input?.samples?.wasm;
  checks.push(
    check(
      'sample-wasm-compile-streaming',
      wasm ?? null,
      'WebAssembly.compileStreaming succeeds from a 200 application/wasm response',
      wasm?.compiled === true &&
        wasm?.status === 200 &&
        normalizedContentType(wasm?.contentType) === 'application/wasm',
    ),
  );

  return { pass: checks.every((item) => item.pass), checks };
}

export function analyzeSecurityProbes(input) {
  const maliciousOrigin = input?.maliciousOrigin;
  const forgedCapability = input?.forgedCapability;
  const replayCapability = input?.replayCapability;
  const pathProbes = Array.isArray(input?.pathProbes) ? input.pathProbes : [];
  const oversizedRange = input?.oversizedRange;
  const rejectedPathProbe = (probe) => {
    if (probe?.bytes !== 0) return false;
    if (probe?.outcome === 'timeout-no-response') return true;
    if (probe?.outcome !== 'ERROR') return false;
    if (probe?.kind === 'syntax') return probe?.errorCode === 'protocol';
    return probe?.kind === 'manifest-boundary' && ['missing', 'protocol'].includes(probe?.errorCode);
  };
  const checks = [
    check(
      'malicious-origin-rejected',
      maliciousOrigin ?? null,
      'no challenge and no connected event',
      maliciousOrigin?.challengeCount === 0 && maliciousOrigin?.connectedCount === 0,
    ),
    check(
      'forged-capability-rejected',
      forgedCapability ?? null,
      'CONNECT_ERROR capability with no connection',
      forgedCapability?.outcome === 'CONNECT_ERROR' &&
        forgedCapability?.errorCode === 'capability' &&
        forgedCapability?.connected === false,
    ),
    check(
      'replayed-capability-rejected',
      replayCapability ?? null,
      'first claim CONNECTED and replay CONNECT_ERROR capability',
      replayCapability?.firstOutcome === 'CONNECTED' &&
        replayCapability?.secondOutcome === 'CONNECT_ERROR' &&
        replayCapability?.secondErrorCode === 'capability',
    ),
    check(
      'path-and-url-probes-rejected',
      pathProbes,
      'at least four canary-backed syntax probes plus one known-object manifest-boundary probe rejected with zero bytes',
      pathProbes.filter((probe) => probe?.kind === 'syntax').length >= 4 &&
        pathProbes.some((probe) => probe?.kind === 'manifest-boundary') &&
        pathProbes.every(rejectedPathProbe),
    ),
    check(
      'oversized-range-rejected',
      oversizedRange ?? null,
      '416 with zero body, Accept-Ranges, and exact unsatisfied Content-Range',
      oversizedRange?.status === 416 &&
        oversizedRange?.bytes === 0 &&
        oversizedRange?.contentLength === 0 &&
        oversizedRange?.acceptRanges === 'bytes' &&
        oversizedRange?.contentRange === oversizedRange?.expectedContentRange,
    ),
  ];
  return { pass: checks.every((item) => item.pass), checks };
}

const TRAFFIC_COUNTER_FIELDS = Object.freeze([
  'workerRequests',
  'cacheHits',
  'r2Heads',
  'r2Gets',
  'declaredBytes',
  'actualBytes',
  'r2Bytes',
  'completed',
  'aborted',
  'failed',
]);
const TRAFFIC_COUNTER_GROUPS = Object.freeze(['segments', 'objects', 'routes']);

function parseTrafficSnapshot(snapshot) {
  if (!snapshot || typeof snapshot !== 'object' || Array.isArray(snapshot)) return null;
  const normalized = {};
  for (const group of TRAFFIC_COUNTER_GROUPS) {
    const counters = snapshot[group];
    if (!counters || typeof counters !== 'object' || Array.isArray(counters)) return null;
    normalized[group] = {};
    for (const [key, counter] of Object.entries(counters)) {
      if (!counter || typeof counter !== 'object' || Array.isArray(counter)) return null;
      const normalizedCounter = {};
      for (const field of TRAFFIC_COUNTER_FIELDS) {
        if (!Number.isSafeInteger(counter[field]) || counter[field] < 0) return null;
        normalizedCounter[field] = counter[field];
      }
      if (!counter.statuses || typeof counter.statuses !== 'object' || Array.isArray(counter.statuses)) return null;
      normalizedCounter.statuses = {};
      for (const [status, count] of Object.entries(counter.statuses)) {
        if (!/^\d{3}$/u.test(status) || !Number.isSafeInteger(count) || count < 0) return null;
        normalizedCounter.statuses[status] = count;
      }
      normalized[group][key] = normalizedCounter;
    }
  }
  return normalized;
}

export function analyzeTrafficEvidence(input) {
  const before = parseTrafficSnapshot(input?.before);
  const after = parseTrafficSnapshot(input?.after);
  const evidenceUrl = input?.evidenceUrl;
  const delta = {};
  const nonzero = [];
  const checks = [
    check(
      'traffic-evidence-interface-present',
      evidenceUrl ?? null,
      'a non-empty HTTP/R2 counter evidence URL',
      typeof evidenceUrl === 'string' && evidenceUrl.length > 0,
    ),
    check('traffic-evidence-before-valid', before, 'a complete nonnegative counter snapshot', before !== null),
    check('traffic-evidence-after-valid', after, 'a complete nonnegative counter snapshot', after !== null),
  ];

  if (before && after) {
    for (const group of TRAFFIC_COUNTER_GROUPS) {
      delta[group] = {};
      const keys = new Set([...Object.keys(before[group]), ...Object.keys(after[group])]);
      for (const key of keys) {
        const beforeCounter = before[group][key];
        const afterCounter = after[group][key];
        if (!beforeCounter || !afterCounter) {
          nonzero.push({
            group,
            key,
            field: 'counter-presence',
            before: beforeCounter ?? null,
            after: afterCounter ?? null,
          });
          continue;
        }
        const counterDelta = {};
        for (const field of TRAFFIC_COUNTER_FIELDS) {
          counterDelta[field] = afterCounter[field] - beforeCounter[field];
          if (counterDelta[field] !== 0) {
            nonzero.push({ group, key, field, delta: counterDelta[field] });
          }
        }
        counterDelta.statuses = {};
        const statuses = new Set([...Object.keys(beforeCounter.statuses), ...Object.keys(afterCounter.statuses)]);
        for (const status of statuses) {
          const statusDelta = (afterCounter.statuses[status] ?? 0) - (beforeCounter.statuses[status] ?? 0);
          counterDelta.statuses[status] = statusDelta;
          if (statusDelta !== 0) {
            nonzero.push({ group, key, field: `statuses.${status}`, delta: statusDelta });
          }
        }
        delta[group][key] = counterDelta;
      }
    }
  }

  checks.push(
    check(
      'zero-worker-and-r2-delta',
      nonzero,
      'all Worker, cache, R2, byte, completion, failure, and status counter deltas equal zero',
      before !== null && after !== null && nonzero.length === 0,
    ),
  );
  return { pass: checks.every((item) => item.pass), checks, delta, nonzero };
}

export function combineAcceptanceSections(sections) {
  const entries = Object.entries(sections);
  return {
    pass: entries.length > 0 && entries.every(([, section]) => section?.pass === true),
    sections: Object.fromEntries(entries),
  };
}
