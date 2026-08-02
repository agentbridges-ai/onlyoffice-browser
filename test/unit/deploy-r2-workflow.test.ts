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
    expect(workflow).toContain('names.size<45');
    expect(workflow).toContain("'Wingdings 3'");
    expect(workflow).toContain('hidden.some(name=>names.has(name))');
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

  it('publishes v5 CAS objects additively and verifies every remote SHA-256 before Worker deployment', () => {
    const workflow = fs.readFileSync(path.resolve('.github/workflows/deploy-r2.yml'), 'utf8');
    const upload = workflow.indexOf('name: Upload additive release and compatibility objects (not verification)');
    const remoteVerification = workflow.indexOf('name: Verify every v5 immutable object directly from R2');
    const worker = workflow.indexOf('name: Deploy wildcard runtime Worker');
    const routeVerification = workflow.indexOf('name: Verify release-specific Worker routes before activation');
    const activation = workflow.indexOf('name: Activate the authoritative v5 stable pointer');
    const postVerification = workflow.indexOf('name: Verify production deployment');

    expect(workflow).toContain('ONLYOFFICE_FASTCDC_EVIDENCE_MODE: automatic');
    expect(workflow).toContain('--fastcdc-evidence-mode "${ONLYOFFICE_FASTCDC_EVIDENCE_MODE}"');
    expect(workflow).toContain('git ls-files --error-unmatch "${ONLYOFFICE_FASTCDC_EVIDENCE_PATH}"');
    expect(workflow).toContain('--package-version "$(node -p "require(\'./package.json\').version")"');
    expect(workflow).toContain('--expected-package-version 0.5.7');
    expect(workflow).toContain('--remote r2:onlyoffice-getpi-work');
    expect(workflow).toContain('partial success is not verification');
    expect(workflow).not.toMatch(/\brclone\s+sync\b/);
    expect(workflow).not.toMatch(/--delete(?:-after|-before|-during)?\b/);
    expect(workflow).not.toMatch(/\brclone\s+check\b/);

    expect(upload).toBeGreaterThanOrEqual(0);
    expect(remoteVerification).toBeGreaterThan(upload);
    expect(worker).toBeGreaterThan(remoteVerification);
    expect(routeVerification).toBeGreaterThan(worker);
    expect(activation).toBeGreaterThan(routeVerification);
    expect(postVerification).toBeGreaterThan(activation);
  });

  it('requires a non-downgradable full-v5 local matrix and retains source-bound JSON evidence', () => {
    const workflow = fs.readFileSync(path.resolve('.github/workflows/deploy-r2.yml'), 'utf8');
    const sourceIdentity = workflow.indexOf('name: Bind release build to source identity');
    const matrix = workflow.indexOf('name: Run mandatory full-v5 local Cloudflare/Broker matrix');
    const evidence = workflow.indexOf('name: Record successful full-v5 matrix evidence');
    const artifact = workflow.indexOf('name: Retain successful full-v5 matrix evidence');
    const upload = workflow.indexOf('name: Upload additive release and compatibility objects (not verification)');
    const activation = workflow.indexOf('name: Activate the authoritative v5 stable pointer');
    const matrixBlock = workflow.slice(matrix, evidence);

    expect(workflow).toContain('ONLYOFFICE_CF_MATRIX_FORCE_FULL: "1"');
    expect(workflow).toContain('ONLYOFFICE_CF_MATRIX_GREP: ""');
    expect(workflow).toContain('ONLYOFFICE_CF_MATRIX_REUSE_BUILD: "0"');
    expect(workflow).toContain('ONLYOFFICE_CF_MATRIX_KEEP_STATE: "0"');
    expect(workflow).not.toContain('ONLYOFFICE_CF_MATRIX_REUSE_BUILD: "1"');
    expect(matrixBlock).toContain('test -z "${ONLYOFFICE_CF_MATRIX_GREP}"');
    expect(matrixBlock).toContain('test "${ONLYOFFICE_CF_MATRIX_REUSE_BUILD}" = "0"');
    expect(matrixBlock).toContain('node scripts/run-cloudflare-local-matrix.mjs');
    expect(matrixBlock).not.toContain('--grep');
    expect(workflow).toContain("execFileSync('git', ['rev-parse', 'HEAD']");
    expect(workflow).toContain('gitCommit !== process.env.GITHUB_SHA');
    expect(workflow).toContain("fs.readFileSync('pnpm-lock.yaml')");
    expect(workflow).toContain("fs.readFileSync('dist/onlyoffice-runtime-assets.json')");
    expect(workflow).toContain('releaseManifest.runtimeManifestSha256 !== runtimeManifestSha256');
    expect(workflow).toContain('release-evidence/source-identity.json');
    expect(workflow).toContain('release-evidence/full-v5-cloudflare-broker-matrix.json');
    expect(workflow).toContain('uses: actions/upload-artifact@v6');
    expect(workflow).toContain('if-no-files-found: error');
    expect(workflow).toContain('retention-days: 90');
    expect(sourceIdentity).toBeGreaterThanOrEqual(0);
    expect(matrix).toBeGreaterThan(sourceIdentity);
    expect(evidence).toBeGreaterThan(matrix);
    expect(artifact).toBeGreaterThan(evidence);
    expect(upload).toBeGreaterThan(artifact);
    expect(activation).toBeGreaterThan(upload);
  });

  it('keeps the legacy v4 pointer frozen and uses stable-v5 as the only activation write', () => {
    const workflow = fs.readFileSync(path.resolve('.github/workflows/deploy-r2.yml'), 'utf8');
    const freezeStart = workflow.indexOf('name: Validate and freeze the legacy v4 compatibility pointer');
    const activationStart = workflow.indexOf('name: Activate the authoritative v5 stable pointer');
    const verificationStart = workflow.indexOf('name: Verify production deployment');
    const freeze = workflow.slice(freezeStart, activationStart);
    const activation = workflow.slice(activationStart, verificationStart);
    const verification = workflow.slice(verificationStart);

    expect(workflow).toContain(
      "if(manifest.version !== 4 || manifest.releaseId !== pointer.releaseId) throw new Error('stable.json must retain the v4 compatibility manifest')",
    );
    expect(workflow).toContain("require('./.onlyoffice-release/channels/stable-v5.json')");
    expect(freezeStart).toBeGreaterThanOrEqual(0);
    expect(activationStart).toBeGreaterThan(freezeStart);
    expect(verificationStart).toBeGreaterThan(activationStart);
    expect(freeze).toContain('rclone_cat_to_file \\');
    expect(freeze).toContain('r2:onlyoffice-getpi-work/channels/stable.json');
    expect(freeze).toContain('timeout 60s rclone cat "${source}"');
    expect(freeze).toContain("'releases/'+pointer.releaseId+'/manifest.json'");
    expect(freeze).toContain('((pointer.manifestUrl === undefined) !== (pointer.manifestSha256 === undefined))');
    expect(freeze).toContain('manifest.version !== 4');
    expect(activation).not.toContain('channels/stable.json');
    expect(activation).toContain('r2:onlyoffice-getpi-work/channels/stable-v5.json');
    expect(activation).toContain('cmp --silent .onlyoffice-release/channels/stable-v5.json -');
    expect(verification).toContain('cmp --silent "${RUNNER_TEMP}/onlyoffice-legacy-stable.json" -');
    expect(verification).toContain("await verify('stable-v5.json', expectedV5, 5, attempt)");
    expect(verification).toContain("await verify('stable.json', expectedLegacy, 4, attempt)");
    expect(workflow).not.toMatch(/rclone_copyto_with_outer_retry[\s\S]*?\.onlyoffice-release\/channels\/stable\.json/);
  });
});
