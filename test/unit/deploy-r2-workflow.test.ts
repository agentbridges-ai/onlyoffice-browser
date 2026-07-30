import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';

describe('R2 release workflow', () => {
  it('hydrates a pinned immutable font release instead of regenerating its own output', () => {
    const workflow = fs.readFileSync(path.resolve('.github/workflows/deploy-r2.yml'), 'utf8');

    expect(workflow).toContain('ONLYOFFICE_MATRIX_FONT_RELEASE_ID: v0.5.0-648c486d1c558acd');
    expect(workflow).toContain(
      'ONLYOFFICE_MATRIX_FONT_MANIFEST_SHA256: 09f0df58c08043aac5b576411b51bfe97584ea5b18e3eac5241b9d448631ba45',
    );
    expect(workflow).toContain('node scripts/hydrate-cloudflare-matrix-fonts.mjs');
    expect(workflow).not.toContain('retained-font-input');
    expect(workflow).not.toContain('--include "/fonts/**"');
  });

  it('derives SOURCE_DATE_EPOCH from the checked-out commit before building', () => {
    const workflow = fs.readFileSync(path.resolve('.github/workflows/deploy-r2.yml'), 'utf8');
    const timestamp = workflow.indexOf('SOURCE_DATE_EPOCH=$(git log -1 --format=%ct)');
    const build = workflow.indexOf('pnpm build');

    expect(timestamp).toBeGreaterThanOrEqual(0);
    expect(build).toBeGreaterThan(timestamp);
  });
});
