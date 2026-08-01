import { IDBFactory } from 'fake-indexeddb';
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  IndexedDbReleaseLeaseLedger,
  MemoryReleaseLeaseLedger,
  RELEASE_GC_MINIMUM_RETENTION_MS,
  RELEASE_GC_EXECUTION_POLICY,
  RELEASE_LEASE_DEFAULT_TTL_MS,
  ReleaseLeaseHeartbeat,
  ReleaseLeaseError,
  conservativeNoDeleteReleaseGcPlan,
  createReleaseLeaseBinding,
  planReleaseGarbageCollection,
  type ReleaseGcInput,
  type ReleaseGcReleaseRecord,
  type ReleaseLeaseBinding,
  type ReleaseLeaseLedgerSnapshot,
} from '../../src/lib/release-lease-gc';

const DAY_MS = 24 * 60 * 60 * 1_000;
const NOW_MS = 250 * DAY_MS;
const ARIES_ORIGIN = 'https://aries.getpi.work' as const;

const digest = (character: string): string => character.repeat(64);

function leaseBinding(overrides: Partial<ReleaseLeaseBinding> = {}): ReleaseLeaseBinding {
  return {
    leaseId: 'lease-0000000001',
    releaseId: 'release-a',
    sessionId: 'session-a',
    editorOrigin: ARIES_ORIGIN,
    ...overrides,
  };
}

function release(
  releaseId: string,
  state: ReleaseGcReleaseRecord['state'],
  ageDays: number,
  objectSha256: string[],
): ReleaseGcReleaseRecord {
  return {
    releaseId,
    state,
    publishedAtMs: NOW_MS - ageDays * DAY_MS,
    objectSha256,
  };
}

function gcInput(releases: ReleaseGcReleaseRecord[], overrides: Partial<ReleaseGcInput> = {}): ReleaseGcInput {
  return {
    nowMs: NOW_MS,
    ledgerState: 'valid',
    releases,
    knownObjectSha256: [...new Set(releases.flatMap((item) => item.objectSha256))],
    leaseLedger: { version: 1, leases: [] },
    piworkDescriptorReleaseIds: [],
    stableReleaseIds: [],
    ...overrides,
  };
}

afterEach(() => {
  vi.useRealTimers();
});

describe('MemoryReleaseLeaseLedger', () => {
  it('binds a short-lived lease to release, session, and editor origin and renews it by heartbeat', () => {
    let nowMs = 1_000;
    const ledger = new MemoryReleaseLeaseLedger({ now: () => nowMs });
    const binding = leaseBinding();

    expect(ledger.acquireLease(binding)).toEqual({
      ...binding,
      acquiredAtMs: 1_000,
      heartbeatAtMs: 1_000,
      expiresAtMs: 1_000 + RELEASE_LEASE_DEFAULT_TTL_MS,
    });

    nowMs = 40_000;
    expect(ledger.heartbeatLease(binding)).toMatchObject({
      heartbeatAtMs: 40_000,
      expiresAtMs: 40_000 + RELEASE_LEASE_DEFAULT_TTL_MS,
    });
    expect(() => ledger.heartbeatLease(leaseBinding({ editorOrigin: 'https://taurus.getpi.work' }))).toThrowError(
      ReleaseLeaseError,
    );

    nowMs = 40_000 + RELEASE_LEASE_DEFAULT_TTL_MS;
    try {
      ledger.heartbeatLease(binding);
      throw new Error('expected heartbeat to reject an expired lease');
    } catch (error) {
      expect(error).toBeInstanceOf(ReleaseLeaseError);
      expect((error as ReleaseLeaseError).code).toBe('expired');
    }
    expect(ledger.reapExpired()).toEqual([binding.leaseId]);
    expect(ledger.listActiveLeases()).toEqual([]);
  });

  it('round-trips deterministic snapshots and rejects malformed or duplicate leases', () => {
    const ledger = new MemoryReleaseLeaseLedger({ now: () => 10_000 });
    ledger.acquireLease(leaseBinding({ leaseId: 'lease-0000000002', sessionId: 'session-b' }));
    ledger.acquireLease(leaseBinding({ leaseId: 'lease-0000000001' }));
    const snapshot = ledger.snapshot();

    expect(snapshot.leases.map((item) => item.leaseId)).toEqual(['lease-0000000001', 'lease-0000000002']);
    expect(new MemoryReleaseLeaseLedger({ snapshot }).snapshot()).toEqual(snapshot);
    expect(
      () =>
        new MemoryReleaseLeaseLedger({
          snapshot: { version: 1, leases: [snapshot.leases[0], snapshot.leases[0]] },
        }),
    ).toThrowError(/duplicate lease/);
  });

  it('recovers an expired binding after a frozen page resumes without weakening binding checks', () => {
    let nowMs = 1_000;
    const ledger = new MemoryReleaseLeaseLedger({ now: () => nowMs, defaultTtlMs: 30 });
    const binding = leaseBinding();
    ledger.acquireLease(binding);

    nowMs = 1_031;
    expect(() => ledger.heartbeatLease(binding)).toThrowError(ReleaseLeaseError);
    expect(ledger.acquireOrRenewLease(binding)).toMatchObject({
      ...binding,
      acquiredAtMs: 1_031,
      heartbeatAtMs: 1_031,
      expiresAtMs: 1_061,
    });
  });
});

describe('IndexedDbReleaseLeaseLedger', () => {
  it('persists leases across instances and atomically preserves concurrent tabs', async () => {
    let nowMs = 1_000;
    const indexedDb = new IDBFactory();
    const databaseName = `lease-ledger-${crypto.randomUUID()}`;
    const first = new IndexedDbReleaseLeaseLedger(indexedDb, {
      databaseName,
      now: () => nowMs,
      defaultTtlMs: 100,
    });
    const second = new IndexedDbReleaseLeaseLedger(indexedDb, {
      databaseName,
      now: () => nowMs,
      defaultTtlMs: 100,
    });
    const aries = leaseBinding({ leaseId: 'lease-aries-000001' });
    const taurus = leaseBinding({
      leaseId: 'lease-taurus-00001',
      sessionId: 'session-b',
      editorOrigin: 'https://taurus.getpi.work',
    });

    await Promise.all([first.acquireLease(aries), second.acquireLease(taurus)]);
    expect((await first.snapshot()).leases.map((lease) => lease.leaseId)).toEqual([aries.leaseId, taurus.leaseId]);

    nowMs = 1_050;
    await Promise.all([second.heartbeatLease(aries), first.heartbeatLease(taurus)]);
    expect(await second.listActiveLeases()).toHaveLength(2);
    await first.releaseLease(aries);
    expect(await second.listActiveLeases()).toEqual([
      expect.objectContaining({
        leaseId: taurus.leaseId,
        releaseId: 'release-a',
      }),
    ]);
  });

  it('recovers after a crash by retaining the old lease until TTL and accepting a new owner lease', async () => {
    let nowMs = 10_000;
    const indexedDb = new IDBFactory();
    const databaseName = `lease-recovery-${crypto.randomUUID()}`;
    const beforeCrash = new IndexedDbReleaseLeaseLedger(indexedDb, {
      databaseName,
      now: () => nowMs,
      defaultTtlMs: 100,
    });
    const oldBinding = leaseBinding({ leaseId: 'lease-before-crash-1' });
    await beforeCrash.acquireLease(oldBinding);

    const afterRefresh = new IndexedDbReleaseLeaseLedger(indexedDb, {
      databaseName,
      now: () => nowMs,
      defaultTtlMs: 100,
    });
    const newBinding = leaseBinding({ leaseId: 'lease-after-refresh-1' });
    await afterRefresh.acquireLease(newBinding);
    expect(await afterRefresh.listActiveLeases()).toHaveLength(2);

    nowMs = 10_101;
    expect(await afterRefresh.reapExpired()).toEqual([newBinding.leaseId, oldBinding.leaseId].sort());
    expect(await beforeCrash.listActiveLeases()).toEqual([]);
  });
});

describe('ReleaseLeaseHeartbeat', () => {
  it('binds browser-native timer receivers before scheduling a lease heartbeat', async () => {
    const originalSetTimeout = globalThis.setTimeout;
    const originalClearTimeout = globalThis.clearTimeout;
    let scheduled = false;
    let cleared = false;
    vi.stubGlobal('setTimeout', function (this: typeof globalThis) {
      if (this !== globalThis) throw new TypeError('Illegal invocation');
      scheduled = true;
      return 1 as unknown as ReturnType<typeof globalThis.setTimeout>;
    });
    vi.stubGlobal('clearTimeout', function (this: typeof globalThis) {
      if (this !== globalThis) throw new TypeError('Illegal invocation');
      cleared = true;
    });
    try {
      const heartbeat = new ReleaseLeaseHeartbeat({
        ledger: new MemoryReleaseLeaseLedger({ now: () => 1_000 }),
        binding: leaseBinding(),
      });
      await heartbeat.start();
      expect(scheduled).toBe(true);
      await heartbeat.stop();
      expect(cleared).toBe(true);
    } finally {
      vi.stubGlobal('setTimeout', originalSetTimeout);
      vi.stubGlobal('clearTimeout', originalClearTimeout);
    }
  });

  it('heartbeats below TTL, reacquires after a long freeze, and leaves crash recovery to TTL', async () => {
    vi.useFakeTimers();
    let nowMs = 1_000;
    const ledger = new MemoryReleaseLeaseLedger({ now: () => nowMs, defaultTtlMs: 30 });
    const binding = createReleaseLeaseBinding(
      {
        releaseId: 'release-a',
        sessionId: 'session-a',
        editorOrigin: ARIES_ORIGIN,
      },
      () => '00000000-0000-4000-8000-000000000001',
    );
    const heartbeat = new ReleaseLeaseHeartbeat({
      ledger,
      binding,
      ttlMs: 30,
      heartbeatIntervalMs: 10,
    });
    await heartbeat.start();

    nowMs = 1_010;
    await vi.advanceTimersByTimeAsync(10);
    expect(ledger.listActiveLeases()[0]).toMatchObject({
      heartbeatAtMs: 1_010,
      expiresAtMs: 1_040,
    });

    nowMs = 1_100;
    await heartbeat.heartbeatNow();
    expect(ledger.listActiveLeases()[0]).toMatchObject({
      acquiredAtMs: 1_100,
      heartbeatAtMs: 1_100,
      expiresAtMs: 1_130,
    });

    await heartbeat.stop({ release: false });
    expect(ledger.snapshot().leases).toHaveLength(1);
    nowMs = 1_131;
    expect(ledger.reapExpired()).toEqual([binding.leaseId]);
  });

  it('uses unique owner leases so closing one overlapping refresh cannot release the other', async () => {
    const ledger = new MemoryReleaseLeaseLedger({ now: () => 1_000 });
    const identity = {
      releaseId: 'release-a',
      sessionId: 'session-a',
      editorOrigin: ARIES_ORIGIN,
    };
    const first = new ReleaseLeaseHeartbeat({
      ledger,
      binding: createReleaseLeaseBinding(identity, () => '00000000-0000-4000-8000-000000000001'),
    });
    const second = new ReleaseLeaseHeartbeat({
      ledger,
      binding: createReleaseLeaseBinding(identity, () => '00000000-0000-4000-8000-000000000002'),
    });
    await Promise.all([first.start(), second.start()]);
    await first.stop();

    expect(ledger.listActiveLeases()).toEqual([
      expect.objectContaining({
        leaseId: second.binding.leaseId,
        releaseId: 'release-a',
      }),
    ]);
    await second.stop();
  });
});

describe('planReleaseGarbageCollection', () => {
  it('keeps physical deletion explicitly disabled until external references are transactionally available', () => {
    expect(RELEASE_GC_EXECUTION_POLICY).toBe('no-delete');
    expect(conservativeNoDeleteReleaseGcPlan()).toEqual({
      blocked: true,
      blockedReason: 'execution-disabled',
      retainedReleaseIds: [],
      retentionMarks: [],
      deleteReleaseIds: [],
      deleteObjectSha256: [],
    });
  });

  it('keeps active A when update B failed and deletes only B-exclusive objects after retention expires', () => {
    const shared = digest('a');
    const failedOnly = digest('b');
    const releases = [
      release('release-a', 'active', 200, [shared]),
      release('release-b', 'failed', 180, [shared, failedOnly]),
      release('release-c', 'failed', 3, [digest('c')]),
      release('release-d', 'failed', 2, [digest('d')]),
      release('release-e', 'failed', 1, [digest('e')]),
    ];

    const plan = planReleaseGarbageCollection(gcInput(releases));

    expect(plan.blocked).toBe(false);
    expect(plan.deleteReleaseIds).toEqual(['release-b']);
    expect(plan.deleteObjectSha256).toEqual([failedOnly]);
    expect(plan.retentionMarks.find((mark) => mark.releaseId === 'release-a')?.reasons).toContain('transaction-active');
    expect(plan.deleteObjectSha256).not.toContain(shared);
  });

  it('protects a rolled-back stable release and independent Piwork descriptor references', () => {
    const releases = [
      release('release-a', 'failed', 220, [digest('a')]),
      release('release-b', 'active', 200, [digest('b')]),
      release('release-piwork', 'failed', 210, [digest('f')]),
      release('release-orphan', 'failed', 190, [digest('0')]),
      release('release-c', 'failed', 3, [digest('c')]),
      release('release-d', 'failed', 2, [digest('d')]),
      release('release-e', 'failed', 1, [digest('e')]),
    ];

    const plan = planReleaseGarbageCollection(
      gcInput(releases, {
        stableReleaseIds: ['release-a'],
        piworkDescriptorReleaseIds: ['release-piwork'],
      }),
    );

    expect(plan.deleteReleaseIds).toEqual(['release-orphan']);
    expect(plan.retentionMarks.find((mark) => mark.releaseId === 'release-a')?.reasons).toEqual(['stable-pointer']);
    expect(plan.retentionMarks.find((mark) => mark.releaseId === 'release-piwork')?.reasons).toEqual([
      'piwork-descriptor',
    ]);
  });

  it('keeps a release while an editor lease is live, then permits collection after expiry', () => {
    const leased = release('release-a', 'failed', 200, [digest('a')]);
    const recent = [
      release('release-c', 'failed', 3, [digest('c')]),
      release('release-d', 'failed', 2, [digest('d')]),
      release('release-e', 'failed', 1, [digest('e')]),
    ];
    const leaseLedger: ReleaseLeaseLedgerSnapshot = {
      version: 1,
      leases: [
        {
          ...leaseBinding(),
          acquiredAtMs: NOW_MS - 60_000,
          heartbeatAtMs: NOW_MS - 20_000,
          expiresAtMs: NOW_MS + 10_000,
        },
      ],
    };

    const livePlan = planReleaseGarbageCollection(gcInput([leased, ...recent], { leaseLedger }));
    expect(livePlan.deleteReleaseIds).toEqual([]);
    expect(livePlan.retentionMarks.find((mark) => mark.releaseId === 'release-a')?.reasons).toEqual([
      'live-editor-lease',
    ]);

    const expiredPlan = planReleaseGarbageCollection(
      gcInput([leased, ...recent], {
        nowMs: NOW_MS + 10_000,
        leaseLedger,
      }),
    );
    expect(expiredPlan.deleteReleaseIds).toEqual(['release-a']);
    expect(expiredPlan.deleteObjectSha256).toContain(digest('a'));
  });

  it('always keeps protected transaction states, the newest three releases, and at least 90 days', () => {
    const exactBoundary: ReleaseGcReleaseRecord = {
      ...release('release-boundary', 'failed', 90, [digest('a')]),
      publishedAtMs: NOW_MS - RELEASE_GC_MINIMUM_RETENTION_MS,
    };
    const releases = [
      release('release-installing', 'installing', 220, [digest('1')]),
      release('release-prepared', 'prepared', 220, [digest('2')]),
      release('release-retained', 'retained', 220, [digest('3')]),
      release('release-old', 'failed', 91, [digest('4')]),
      exactBoundary,
      release('release-c', 'failed', 3, [digest('c')]),
      release('release-d', 'failed', 2, [digest('d')]),
      release('release-e', 'failed', 1, [digest('e')]),
    ];

    const plan = planReleaseGarbageCollection(gcInput(releases));

    expect(plan.deleteReleaseIds).toEqual(['release-old']);
    expect(plan.retainedReleaseIds).toEqual([
      'release-boundary',
      'release-c',
      'release-d',
      'release-e',
      'release-installing',
      'release-prepared',
      'release-retained',
    ]);
  });

  it('deletes a shared digest only after every referencing release is collectible', () => {
    const shared = digest('a');
    const releases = [
      release('release-a', 'failed', 220, [shared]),
      release('release-b', 'failed', 210, [shared]),
      release('release-c', 'failed', 3, [digest('c')]),
      release('release-d', 'failed', 2, [digest('d')]),
      release('release-e', 'failed', 1, [digest('e')]),
    ];
    const allCollectible = planReleaseGarbageCollection(gcInput(releases));
    expect(allCollectible.deleteReleaseIds).toEqual(['release-a', 'release-b']);
    expect(allCollectible.deleteObjectSha256).toContain(shared);

    const oneReferenced = planReleaseGarbageCollection(
      gcInput(releases, { piworkDescriptorReleaseIds: ['release-a'] }),
    );
    expect(oneReferenced.deleteReleaseIds).toEqual(['release-b']);
    expect(oneReferenced.deleteObjectSha256).not.toContain(shared);
  });

  it('fails closed with empty deletion sets for unknown, corrupt, or internally inconsistent ledgers', () => {
    const valid = gcInput([
      release('release-a', 'failed', 220, [digest('a')]),
      release('release-c', 'failed', 3, [digest('c')]),
      release('release-d', 'failed', 2, [digest('d')]),
      release('release-e', 'failed', 1, [digest('e')]),
    ]);

    expect(planReleaseGarbageCollection({ ...valid, ledgerState: 'unknown' })).toMatchObject({
      blocked: true,
      blockedReason: 'ledger-unknown',
      deleteReleaseIds: [],
      deleteObjectSha256: [],
    });
    expect(planReleaseGarbageCollection({ ...valid, ledgerState: 'corrupt' })).toMatchObject({
      blocked: true,
      blockedReason: 'ledger-corrupt',
      deleteReleaseIds: [],
      deleteObjectSha256: [],
    });
    expect(
      planReleaseGarbageCollection({
        ...valid,
        piworkDescriptorReleaseIds: ['missing-release'],
      }),
    ).toMatchObject({
      blocked: true,
      blockedReason: 'ledger-corrupt',
      deleteReleaseIds: [],
      deleteObjectSha256: [],
    });
    expect(
      planReleaseGarbageCollection({
        ...valid,
        knownObjectSha256: [digest('c'), digest('d'), digest('e')],
      }),
    ).toMatchObject({
      blocked: true,
      blockedReason: 'ledger-corrupt',
      deleteReleaseIds: [],
      deleteObjectSha256: [],
    });
  });

  it('produces the same sorted plan regardless of ledger iteration order', () => {
    const releases = [
      release('release-a', 'failed', 220, [digest('a')]),
      release('release-b', 'failed', 210, [digest('b')]),
      release('release-c', 'failed', 3, [digest('c')]),
      release('release-d', 'failed', 2, [digest('d')]),
      release('release-e', 'failed', 1, [digest('e')]),
    ];
    const forward = gcInput(releases);
    const reverse = gcInput([...releases].reverse(), {
      knownObjectSha256: [...forward.knownObjectSha256].reverse(),
    });

    expect(planReleaseGarbageCollection(reverse)).toEqual(planReleaseGarbageCollection(forward));
  });
});
