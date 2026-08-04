import fs from 'node:fs';
import { describe, expect, it } from 'vitest';

const audit = fs.readFileSync('.github/workflows/audit-r2.yml', 'utf8');

describe('published runtime audit workflow', () => {
  it('always performs a substantive public HTTP audit first', () => {
    const publicAudit = audit.slice(
      audit.indexOf('Audit the public HTTP publication'),
      audit.indexOf('Report degraded public-only audit'),
    );
    expect(publicAudit).toContain('https://onlyoffice.getpi.work');
    expect(publicAudit).toContain('https://office-editor-github-actions-smoke.getpi.work');
    expect(publicAudit).toContain('public stable-v5 pointer is not served with no-store');
    expect(publicAudit).toContain('sha256(manifestBytes) !== pointer.manifestSha256');
    expect(publicAudit).toContain('sha256(objectBytes) !== asset.sha256');
    expect(publicAudit).toContain('Public HTTP audit: passed');
    expect(publicAudit).not.toContain('secrets.');
    expect(publicAudit).not.toMatch(/\brclone\b/);
  });

  it('reports degraded mode when audit-only storage credentials are absent', () => {
    expect(audit).toContain("if: steps.r2-audit.outputs.enabled != 'true'");
    expect(audit).toContain('Degraded storage audit');
    expect(audit).toContain('complete storage set and protected release/Worker pair authorization were not verified');
    expect(audit).toContain('degraded public-only mode');
    expect(audit).toContain('protected release/Worker pair authorization were not verified');
    expect(audit).toContain('Release/Worker pair authorization: skipped');
  });

  it('uses only optional audit-environment read-only credentials for a full remote re-hash', () => {
    expect(audit).toContain('environment: audit');
    expect(audit).toContain('secrets.R2_AUDIT_READONLY_ACCESS_KEY_ID');
    expect(audit).toContain('secrets.R2_AUDIT_READONLY_SECRET_ACCESS_KEY');
    expect(audit).not.toContain('${{ secrets.R2_ACCESS_KEY_ID }}');
    expect(audit).not.toContain('${{ secrets.R2_SECRET_ACCESS_KEY }}');
    expect(audit).toContain("if: steps.r2-audit.outputs.enabled == 'true'");
    expect(audit).toContain('--v5-manifest manifest-public.json');
    expect(audit).toContain('--remote-verification-mode full');
    expect(audit).toContain('--concurrency 16');
    expect(audit).toContain('cmp stable-v5-public.json stable-v5-r2.json');
    expect(audit).toContain('cmp manifest-public.json manifest-r2.json');
    expect(audit).toContain('permissions:\n  contents: read\n  actions: read');
    expect(audit).toContain('promotions/${RELEASE_ID}');
    expect(audit).toContain('rollbacks/${RELEASE_ID}');
    expect(audit).toContain('promotion receipt digest does not match its immutable object name');
    expect(audit).toContain('rollback receipt digest does not match its immutable object name');
    expect(audit).toContain('receipt.worker.candidateVersionId===process.env.WORKER_VERSION_ID');
    expect(audit).toContain('receipt.worker.versionId===process.env.WORKER_VERSION_ID');
    expect(audit).toContain('/actions/runs/${receipt_run_id}/attempts/${receipt_run_attempt}');
    expect(audit).toContain("'.github/workflows/deploy-r2.yml'");
    expect(audit).toContain("'.github/workflows/rollback-r2.yml'");
    expect(audit).toContain("run.conclusion==='success'");
    expect(audit).toContain('no immutable promotion or promotion-chained rollback receipt authorizes');
    expect(audit).toContain('Release/Worker pair authorization: passed');
    expect(audit).toContain('!positive(receipt.candidate?.runId)');
    expect(audit).toContain('!positive(receipt.staging?.runId)');
    expect(audit).toContain('!positive(receipt.piwork?.deepVerifyRunId)');
    expect(audit).toContain('!positive(receipt.piwork?.deepVerifyRunAttempt)');
    expect(audit).toContain("const route=selected.key.split('/').map(encodeURIComponent).join('/')");
    expect(audit).toContain("response.headers.get('x-onlyoffice-worker-version')!==process.env.WORKER_VERSION_ID");
    expect(audit).toContain('public release-pair receipt bytes, digest, immutable cache policy');
  });

  it('never mutates R2 and pins every third-party action', () => {
    expect(audit).not.toMatch(/rclone (?:copy|copyto) [^\n]+ r2:onlyoffice-getpi-work/);
    expect(audit).not.toMatch(/uses:\s+[^\s]+@v\d/);
  });
});
