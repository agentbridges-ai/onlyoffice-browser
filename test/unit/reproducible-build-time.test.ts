import { describe, expect, it } from 'vitest';

// Executable build helper; intentionally shipped as ESM rather than compiled TypeScript.
// @ts-expect-error JavaScript build script has no declaration output.
import { reproducibleBuildTimestamp } from '../../scripts/reproducible-build-time.mjs';

describe('reproducible build timestamp', () => {
  it('uses SOURCE_DATE_EPOCH so branch and tag builds of one commit are byte-identical', () => {
    expect(reproducibleBuildTimestamp({ SOURCE_DATE_EPOCH: '1785384000' })).toBe('2026-07-30T04:00:00.000Z');
  });

  it('uses the clock only when a reproducible timestamp was not supplied', () => {
    expect(reproducibleBuildTimestamp({}, () => new Date('2026-07-30T05:00:00.000Z'))).toBe('2026-07-30T05:00:00.000Z');
  });

  it('rejects malformed and unsafe values', () => {
    expect(() => reproducibleBuildTimestamp({ SOURCE_DATE_EPOCH: '-1' })).toThrow('non-negative integer');
    expect(() => reproducibleBuildTimestamp({ SOURCE_DATE_EPOCH: 'not-a-time' })).toThrow('non-negative integer');
    expect(() => reproducibleBuildTimestamp({ SOURCE_DATE_EPOCH: '999999999999999999999' })).toThrow(
      'outside the supported',
    );
  });
});
