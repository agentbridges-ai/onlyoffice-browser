import { describe, expect, it } from 'vitest';
import {
  analyzeBrokerMetrics,
  analyzeLifecycleReport,
  analyzeRecoveryResults,
  analyzeReleaseIntegrity,
  analyzeSecurityProbes,
  analyzeStartupPerformance,
  analyzeTrafficEvidence,
  combineAcceptanceSections,
  linearSlope,
  percentile,
} from '../../scripts/canonical-broker-acceptance-policy.mjs';

function lifecycleReport(iterations = 100) {
  const samples = [
    { phase: 'baseline', iteration: 0, readyWaitMs: null as number | null, jsHeapUsedMb: 20, browserRssMb: 200 },
    ...Array.from({ length: iterations }, (_, index) => ({
      phase: 'closed',
      iteration: index + 1,
      readyWaitMs: 220 + (index % 5),
      jsHeapUsedMb: 24 + (index % 3) * 0.1,
      browserRssMb: 230 + (index % 4),
    })),
  ];
  return {
    samples,
    analysis: {
      closedHeapDeltaMb: 4.2,
      closedRssDeltaMb: 33,
      messagePortMetricsAvailable: true,
      finalMessagePortsClean: true,
      finalPortState: {
        activeHostPortCount: 0,
        activeStartupHeartbeatPortCount: 0,
        activeInstanceCount: 0,
        activeOriginLeaseCount: 0,
      },
      unexpectedPageRefreshes: [],
    },
  };
}

describe('canonical Broker production acceptance policy', () => {
  it('uses deterministic nearest-rank P95 and linear slopes', () => {
    expect(percentile([5, 1, 4, 3, 2], 0.95)).toBe(5);
    expect(linearSlope([1, 2, 3, 4])).toBe(1);
  });

  it('passes only a complete 100-cycle lifecycle with reclaimed ports and a stable tail', () => {
    expect(analyzeLifecycleReport(lifecycleReport()).pass).toBe(true);
    expect(analyzeLifecycleReport(lifecycleReport(99)).pass).toBe(false);
    const missingMetrics = lifecycleReport();
    missingMetrics.analysis.messagePortMetricsAvailable = false;
    expect(analyzeLifecycleReport(missingMetrics).pass).toBe(false);
  });

  it('fixes the startup threshold before evaluation as max(300 ms, baseline 10%)', () => {
    const report = lifecycleReport();
    const result = analyzeStartupPerformance(report, 100);
    expect(result).toMatchObject({
      pass: true,
      baselineP95Ms: 100,
      allowedExtraMs: 300,
    });
    report.samples
      .filter((sample) => sample.phase === 'closed')
      .forEach((sample) => {
        sample.readyWaitMs = 450;
      });
    expect(analyzeStartupPerformance(report, 100).pass).toBe(false);
  });

  it('requires three real editor Service Worker snapshots and enforces the 64 MiB high-water cap', () => {
    const canonical = {
      role: 'canonical-service-worker',
      canonical: { activeReads: 0, reservedBytes: 0, peakReservedBytes: 48 * 1024 * 1024 },
    };
    const editor = (peakReservedBytes: number) => ({
      role: 'editor-service-worker',
      editor: { activeRequests: 0, reservedBytes: 0, peakReservedBytes },
    });
    const relay = (peakReservedBytes: number) => ({
      role: 'canonical-relay-frame',
      relay: { activeRequests: 0, reservedBytes: 0, peakReservedBytes },
    });
    const input = {
      snapshots: [canonical, editor(1024), editor(2048), editor(4096), relay(1024), relay(2048), relay(4096)],
      finalEditorSnapshots: [
        { connectionStatus: 'disconnected', activeRequests: 0, reservedBytes: 0 },
        { connectionStatus: 'disconnected', activeRequests: 0, reservedBytes: 0 },
        { connectionStatus: 'disconnected', activeRequests: 0, reservedBytes: 0 },
      ],
      burstStatus: { ready: 3, uniqueOrigins: 3, errors: [] },
    };
    expect(analyzeBrokerMetrics(input).pass).toBe(true);
    input.finalEditorSnapshots[0].connectionStatus = 'terminated';
    expect(analyzeBrokerMetrics(input).pass).toBe(true);
    canonical.canonical.peakReservedBytes = 64 * 1024 * 1024;
    expect(analyzeBrokerMetrics(input).pass).toBe(false);
  });

  it('requires every recovery case and rejects missing structured error evidence', () => {
    const results = [
      { name: 'broker-iframe', outcome: 'recovered', elapsedMs: 100 },
      { name: 'canonical-service-worker', outcome: 'recovered', elapsedMs: 200 },
      { name: 'editor-service-worker', outcome: 'accurate-error', elapsedMs: 30_000, errorCode: 'timeout' },
      { name: 'message-port', outcome: 'recovered', elapsedMs: 100 },
      { name: 'freeze-resume', outcome: 'recovered', elapsedMs: 100 },
    ];
    const recovery = analyzeRecoveryResults(results);
    expect(recovery.pass).toBe(true);
    expect(combineAcceptanceSections({ lifecycle: analyzeLifecycleReport(lifecycleReport()), recovery }).pass).toBe(
      true,
    );
    expect(analyzeRecoveryResults(results.slice(1)).pass).toBe(false);
  });

  it('requires every manifest asset and exact Broker byte, MIME, Range, cancellation, and WASM evidence', () => {
    const digest = 'a'.repeat(64);
    const asset = {
      path: 'sdkjs/app.js',
      expectedBytes: 4,
      expectedSha256: digest,
      expectedMime: 'text/javascript',
      status: 200,
      bytes: 4,
      bytesSent: 4,
      chunks: 1,
      sha256: digest,
      contentType: 'text/javascript; charset=utf-8',
      contentLength: 4,
      acceptRanges: 'bytes',
      terminalType: 'END',
    };
    const input = {
      expectedAssetCount: 1,
      assets: [asset],
      samples: {
        full: asset,
        range: {
          path: 'sdkjs/app.js',
          status: 206,
          start: 1,
          end: 3,
          bytes: 3,
          contentLength: 3,
          contentRange: 'bytes 1-3/4',
          expectedContentRange: 'bytes 1-3/4',
          sha256: digest,
          directSha256: digest,
          expectedSha256: digest,
          expectedSource: 'full-broker-stream-capture',
          contentType: 'text/javascript',
          expectedMime: 'text/javascript',
          acceptRanges: 'bytes',
          terminalType: 'END',
        },
        unsatisfied: {
          status: 416,
          bytes: 0,
          bytesSent: 0,
          contentLength: 0,
          contentRange: 'bytes */4',
          expectedContentRange: 'bytes */4',
          acceptRanges: 'bytes',
          terminalType: 'END',
        },
        cancellation: {
          status: 206,
          terminalType: 'CANCELLED',
          bytes: 1,
          bytesSent: 1,
          expectedBytes: 4,
          elapsedMs: 10,
          editor: {
            status: 206,
            contentLength: 4,
            contentRange: 'bytes 0-3/4',
            expectedContentRange: 'bytes 0-3/4',
            acceptRanges: 'bytes',
            firstBytes: 1,
            firstDone: false,
            aborted: true,
            errorName: 'AbortError',
            elapsedMs: 10,
          },
          drain: {
            drained: true,
            elapsedMs: 10,
            canonical: { activeReads: 0, reservedBytes: 0 },
            relay: { activeRequests: 0, reservedBytes: 0 },
            editor: { activeRequests: 0, reservedBytes: 0 },
          },
          protocolRecovery: {
            status: 206,
            bytes: 1,
            bytesSent: 1,
            contentLength: 1,
            contentRange: 'bytes 0-0/4',
            expectedContentRange: 'bytes 0-0/4',
            terminalType: 'END',
          },
          recovery: {
            status: 206,
            bytes: 1,
            contentLength: 1,
            contentRange: 'bytes 0-0/4',
            expectedContentRange: 'bytes 0-0/4',
            acceptRanges: 'bytes',
            sha256: digest,
            expectedSha256: digest,
            expectedSource: 'full-broker-stream-capture',
            terminalType: 'END',
          },
        },
        wasm: { compiled: true, status: 200, contentType: 'application/wasm' },
      },
    };
    expect(analyzeReleaseIntegrity(input).pass).toBe(true);
    (input.samples.cancellation.drain as { relay: { activeRequests: number; reservedBytes: number } | null }).relay =
      null;
    expect(analyzeReleaseIntegrity(input).pass).toBe(true);
    (input.samples.cancellation.drain as { relay: { activeRequests: number; reservedBytes: number } | null }).relay = {
      activeRequests: 1,
      reservedBytes: 1,
    };
    expect(analyzeReleaseIntegrity(input).pass).toBe(false);
    (input.samples.cancellation.drain as { relay: { activeRequests: number; reservedBytes: number } | null }).relay = {
      activeRequests: 0,
      reservedBytes: 0,
    };
    asset.sha256 = 'b'.repeat(64);
    expect(analyzeReleaseIntegrity(input).pass).toBe(false);
  });

  it('requires malicious-origin, forged/replayed capability, path, and oversized Range rejection', () => {
    const input = {
      maliciousOrigin: { challengeCount: 0, connectedCount: 0 },
      forgedCapability: { outcome: 'CONNECT_ERROR', errorCode: 'capability', connected: false },
      replayCapability: {
        firstOutcome: 'CONNECTED',
        secondOutcome: 'CONNECT_ERROR',
        secondErrorCode: 'capability',
      },
      pathProbes: [
        { kind: 'syntax', outcome: 'timeout-no-response', bytes: 0 },
        { kind: 'syntax', outcome: 'ERROR', errorCode: 'protocol', bytes: 0 },
        { kind: 'syntax', outcome: 'ERROR', errorCode: 'protocol', bytes: 0 },
        { kind: 'syntax', outcome: 'timeout-no-response', bytes: 0 },
        { kind: 'manifest-boundary', outcome: 'ERROR', errorCode: 'missing', bytes: 0 },
      ],
      oversizedRange: {
        status: 416,
        bytes: 0,
        contentLength: 0,
        acceptRanges: 'bytes',
        contentRange: 'bytes */4',
        expectedContentRange: 'bytes */4',
      },
    };
    expect(analyzeSecurityProbes(input).pass).toBe(true);
    input.pathProbes[0].bytes = 1;
    expect(analyzeSecurityProbes(input).pass).toBe(false);
  });

  it('fails closed without complete HTTP/R2 evidence and rejects content traffic while retaining shell activity', () => {
    const counter = {
      workerRequests: 2,
      cacheHits: 1,
      r2Heads: 0,
      r2Gets: 0,
      declaredBytes: 128,
      actualBytes: 128,
      r2Bytes: 0,
      completed: 2,
      aborted: 0,
      failed: 0,
      statuses: { '200': 2 },
    };
    const before = {
      segments: { all: structuredClone(counter) },
      objects: {},
      routes: { all: structuredClone(counter) },
    };
    expect(
      analyzeTrafficEvidence({ evidenceUrl: 'https://metrics.example/counters', before, after: before }).pass,
    ).toBe(true);
    const shellActivity = structuredClone(before);
    shellActivity.routes.all.workerRequests += 1;
    shellActivity.routes.all.completed += 1;
    shellActivity.routes.all.statuses['200'] += 1;
    const shellResult = analyzeTrafficEvidence({
      evidenceUrl: 'https://metrics.example/counters',
      before,
      after: shellActivity,
    });
    expect(shellResult.pass).toBe(true);
    expect(shellResult.routeActivity).toHaveLength(3);
    const after = structuredClone(before);
    after.segments.all.r2Gets += 1;
    expect(analyzeTrafficEvidence({ evidenceUrl: 'https://metrics.example/counters', before, after }).pass).toBe(false);
    expect(analyzeTrafficEvidence({ evidenceUrl: '', before, after: before }).pass).toBe(false);
    expect(
      analyzeTrafficEvidence({ evidenceUrl: 'https://metrics.example/counters', before: null, after: null }).pass,
    ).toBe(false);
  });
});
