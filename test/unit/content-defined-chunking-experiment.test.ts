import { describe, expect, it } from 'vitest';
// @ts-expect-error Experimental Node script intentionally has no declaration output.
import { compareSyntheticScenario } from '../../scripts/experiment-content-defined-chunking.mjs';

function bytes(length: number, value: number): Buffer {
  return Buffer.alloc(length, value);
}

describe('content-defined chunking experiment harness', () => {
  it('shows why a prefix insertion invalidates fixed offsets but not file-level CAS', () => {
    const base = [
      { name: 'b', bytes: bytes(12, 1) },
      { name: 'c', bytes: bytes(12, 2) },
    ];
    const inserted = { name: 'a', bytes: bytes(3, 3) };
    const result = compareSyntheticScenario(
      {
        name: 'prefix-insert',
        from: base,
        to: [inserted, ...base],
        actualChangedBytes: inserted.bytes.byteLength,
      },
      undefined,
    );

    expect(result.fixed.downloadBytes).toBeGreaterThan(result.actualChangedBytes);
    expect(result.fileCas.downloadBytes).toBe(result.actualChangedBytes);
  });

  it('reuses every file object when only manifest order changes', () => {
    const first = { name: 'a', bytes: bytes(12, 1) };
    const second = { name: 'b', bytes: bytes(12, 2) };
    const result = compareSyntheticScenario(
      {
        name: 'file-reorder',
        from: [first, second],
        to: [second, first],
        actualChangedBytes: 0,
      },
      undefined,
    );

    expect(result.fileCas.downloadBytes).toBe(0);
    expect(result.fileCas.reusedObjects).toBe(2);
  });
});
