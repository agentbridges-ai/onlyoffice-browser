import fs from 'node:fs';
import { describe, expect, it } from 'vitest';

const stage = fs.readFileSync('.github/workflows/stage-r2.yml', 'utf8');

describe('immutable candidate staging workflow', () => {
  it('stages and fully verifies only the retained candidate CAS against the current Worker', () => {
    expect(stage).toContain("inputs.confirm_stage == 'STAGE'");
    expect(stage).toContain('candidate_commit');
    expect(stage).toContain('candidate_run_id');
    expect(stage).toContain('environment: production');
    expect(stage).toContain('timeout-minutes: 180');
    expect(stage).toContain('onlyoffice-runtime-candidate-${{ inputs.candidate_commit }}');
    expect(stage).toContain("r.path.endsWith('/candidate-r2.yml')");
    expect(stage).toContain('validateReleaseEnvelope');
    expect(stage).toContain('bindExistingNpmPublication');
    expect(stage).toContain('candidate npm package must be published with verified provenance before staging');
    expect(stage).toContain('new npm publication is not the exact candidate commit and source package');
    expect(stage).toContain("envelope.npm.registry.state==='published'");
    expect(stage).toContain('verified-npm-registry.json');
    expect(stage).toContain('npmRegistry');
    expect(stage).toContain('candidate matrix evidence mismatch');
    expect(stage).toContain('manifest.sourceCommit!==process.env.CANDIDATE');
    expect(stage).toContain('r2:onlyoffice-getpi-work/blobs --ignore-existing');
    expect(stage).toContain('r2:onlyoffice-getpi-work/releases --ignore-existing');
    expect(stage).toContain('r2:onlyoffice-getpi-work/packages --ignore-existing');
    expect(stage).toContain('r2:onlyoffice-getpi-work/segments --ignore-existing');
    expect(stage).toContain('--remote-verification-mode full');
    expect(stage).toContain('Verify staged release against the current production Worker');
    expect(stage).toContain('scripts/verify-release-http.mjs');
    expect(stage).toContain('current-production-worker-compatibility');
    expect(stage).toContain('Retain staged compatibility evidence');
    expect(stage).not.toMatch(/channels\/(?:stable|stable-v5)\.json/);
    expect(stage).not.toMatch(/rclone copy candidate-input\/dist/);
    expect(stage).not.toContain('wrangler');
    expect(stage).not.toContain('CLOUDFLARE_API_TOKEN');
  });

  it('pins every third-party action to a full commit SHA', () => {
    expect(stage).not.toMatch(/uses:\s+[^\s]+@v\d/);
  });
});
