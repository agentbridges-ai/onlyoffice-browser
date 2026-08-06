import { describe, expect, it } from 'vitest';
import {
  assertReleaseStateHistory,
  assertReleaseTransition,
  releaseStateEvidence,
  // @ts-expect-error JavaScript release helper intentionally has no declaration output.
} from '../../scripts/release-state-machine.mjs';

describe('release state machine', () => {
  it('accepts only the ordered idempotent promotion lifecycle', () => {
    const history = ['candidate', 'staged', 'verified', 'canary', 'promoted', 'post-verified'];
    expect(assertReleaseStateHistory(history)).toEqual(history);
    expect(releaseStateEvidence(history, { releaseId: 'v0.5.15-test' })).toMatchObject({
      state: 'post-verified',
      releaseId: 'v0.5.15-test',
    });
  });

  it('rejects skipped gates and rollback-shaped forward transitions', () => {
    expect(() => assertReleaseTransition('candidate', 'verified')).toThrow(/invalid release transition/);
    expect(() => assertReleaseStateHistory(['candidate', 'staged', 'canary'])).toThrow(/invalid release transition/);
    expect(() => assertReleaseStateHistory(['candidate', 'staged', 'verified', 'canary', 'post-verified'])).toThrow(
      /invalid release transition/,
    );
  });
});
