import fs from 'node:fs';
import { describe, expect, it } from 'vitest';

const runner = fs.readFileSync('scripts/run-cloudflare-local-matrix.mjs', 'utf8');

describe('local Cloudflare matrix runner startup', () => {
  it('allows cold Wrangler/R2 startup without masking a real readiness failure', () => {
    expect(runner).toContain('ONLYOFFICE_CF_MATRIX_READY_TIMEOUT_MS');
    expect(runner).toContain("|| '180000'");
    expect(runner).toContain('readyTimeoutMs < 30_000');
    expect(runner).toContain('readyTimeoutMs > 300_000');
    expect(runner).toContain('within ${readyTimeoutMs}ms');
  });
});
