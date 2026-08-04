import fs from 'node:fs';
import { describe, expect, it } from 'vitest';

const candidate = fs.readFileSync('.github/workflows/candidate-r2.yml', 'utf8');
const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8')) as {
  scripts?: Record<string, string>;
};

describe('immutable candidate workflow', () => {
  it('has no deployment environment, secret or writable R2 dependency', () => {
    expect(candidate).not.toContain('environment:');
    expect(candidate).not.toContain('secrets.');
    expect(candidate).not.toMatch(/\brclone\b/);
    expect(candidate).not.toContain('r2:onlyoffice-getpi-work');
    expect(candidate).not.toContain('office-editor-');
    expect(candidate).not.toContain('verify-release-http');
    expect(candidate).not.toContain('--remote');
  });

  it('restores or read-only hydrates and then saves the digest-pinned font input', () => {
    expect(candidate).toContain('actions/cache/restore@0057852bfaa89a56745cba8c7296529d2fc39830');
    expect(candidate).toContain('actions/cache/save@0057852bfaa89a56745cba8c7296529d2fc39830');
    expect(candidate).not.toContain('fail-on-cache-miss: true');
    expect(candidate).toContain('steps.full-font-cache.outputs.cache-hit');
    expect(candidate).toContain('ONLYOFFICE_MATRIX_FONT_MANIFEST_SHA256');
    expect(candidate).toContain('shasum -a 256');
    expect(candidate).toContain('ONLYOFFICE_MATRIX_FONT_ASSETS_DIR=');
    expect(candidate).toContain('ONLYOFFICE_MATRIX_FONT_SOURCE=https://onlyoffice.getpi.work');
    expect(candidate).toContain('ONLYOFFICE_MATRIX_FONT_CACHE=');
    expect(candidate).toContain("if: steps.full-font-cache.outputs.cache-hit != 'true'");
  });

  it('verifies the exact public x2t release, source commit, and attested local payload before building', () => {
    expect(candidate).toContain('node scripts/verify-x2t-publication.mjs');
    expect(candidate).toContain('--version 9.3.0+3');
    expect(candidate).toContain('--tag v9.3.0+3');
    expect(candidate).toContain('--commit 4360541a805477726a06caabd1c64a330377bcf2');
    expect(candidate).toContain('--asset-root public/wasm/x2t');
    expect(candidate.indexOf('Verify the pinned public x2t publication and provenance')).toBeLessThan(
      candidate.indexOf('Test and build the candidate exactly once'),
    );
  });

  it('builds once, runs both local matrices and verifies the local release', () => {
    expect(candidate.match(/ONLYOFFICE_SKIP_RELEASE_BUILD=1 pnpm build/g)).toHaveLength(1);
    expect(candidate.match(/node scripts\/build-release-manifest\.mjs/g)).toHaveLength(1);
    expect(candidate).toContain('--source-commit "${GITHUB_SHA}"');
    expect(candidate).toContain('release manifest does not bind GITHUB_SHA');
    expect(candidate).toContain('node scripts/verify-release-publication.mjs');
    expect(candidate).toContain('Run mandatory full-v5 Cloudflare and Broker matrix');
    expect(candidate).toContain('ONLYOFFICE_CF_MATRIX_USE_CURRENT_BUILD=1');
    expect(candidate).toContain('Run fault-injected proxy matrix');
    expect(candidate).toContain("ONLYOFFICE_CF_MATRIX_GREP='fault-injected segment stream aborts'");
    expect(candidate.indexOf('Test and build the candidate exactly once')).toBeLessThan(
      candidate.indexOf('Run mandatory full-v5 Cloudflare and Broker matrix'),
    );
    expect(candidate.indexOf('Run mandatory full-v5 Cloudflare and Broker matrix')).toBeLessThan(
      candidate.indexOf('Run fault-injected proxy matrix'),
    );
  });

  it('retains one GitHub artifact as the sole promotion input after every gate', () => {
    expect(candidate).toContain('timeout-minutes: 180');
    expect(candidate).toContain('Inventory candidate artifact and enforce a conservative size ceiling');
    expect(candidate).toContain('candidate-artifact-inventory.json');
    expect(candidate).toContain('const policyMaxBytes = 4 * 1024 ** 3');
    expect(candidate).toContain('candidate artifact must not contain symlinks');
    expect(candidate).toContain('steps.artifact-inventory.outputs.files');
    expect(candidate).toContain('steps.artifact-inventory.outputs.bytes');
    expect(candidate.match(/actions\/upload-artifact@/g)).toHaveLength(1);
    expect(candidate).toContain('name: onlyoffice-runtime-candidate-${{ github.sha }}');
    expect(candidate).toContain('id: candidate-artifact');
    expect(candidate).toContain('artifact-digest');
    expect(candidate).toContain('retention-days: 30');
    expect(candidate).toContain('dist\n            .onlyoffice-release\n            release-evidence');
    expect(candidate.indexOf('Run fault-injected proxy matrix')).toBeLessThan(
      candidate.indexOf('Inventory candidate artifact and enforce a conservative size ceiling'),
    );
    expect(candidate.indexOf('Inventory candidate artifact and enforce a conservative size ceiling')).toBeLessThan(
      candidate.indexOf('Retain the sole fully verified promotion input'),
    );
  });

  it('does not publish an npm script whose implementation is excluded from the package', () => {
    expect(packageJson.scripts).not.toHaveProperty('release:envelope');
  });
});
