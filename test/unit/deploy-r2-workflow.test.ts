import fs from 'node:fs';
import { describe, expect, it } from 'vitest';

const read = (file: string) => fs.readFileSync(file, 'utf8');
const candidate = read('.github/workflows/candidate-r2.yml');
const promotion = read('.github/workflows/deploy-r2.yml');
const rollback = read('.github/workflows/rollback-r2.yml');
const audit = read('.github/workflows/audit-r2.yml');
const npmPublication = read('.github/workflows/release-npm.yml');
const pullRequestChecks = read('.github/workflows/ci.yml');
const matrixRunner = read('scripts/run-cloudflare-local-matrix.mjs');

describe('R2 release train workflows', () => {
  it('builds, matrices, publishes immutable CAS, smokes, then retains a candidate', () => {
    expect(candidate).toContain('SOURCE_DATE_EPOCH=$(git log -1 --format=%ct)');
    expect(candidate).toContain("if: github.ref == 'refs/heads/main'");
    expect(candidate).toContain('ONLYOFFICE_SKIP_RELEASE_BUILD=1 pnpm build');
    expect(candidate).toContain('Record immutable source identity before matrices');
    expect(candidate).toContain('ONLYOFFICE_CF_MATRIX_USE_CURRENT_BUILD=1');
    expect(candidate).toContain('matrix changed candidate source identity');
    expect(candidate).toContain('fault matrix changed candidate source identity');
    expect(candidate).toContain('Run mandatory full-v5 Cloudflare and Broker matrix');
    expect(candidate).toContain('Run fault-injected proxy matrix');
    expect(candidate).not.toContain('r2:onlyoffice-getpi-work');
    expect(candidate).toContain('Retain the sole fully verified promotion input');
    expect(candidate).not.toContain('Smoke the exact candidate route');
    expect(candidate.indexOf('Test and build the candidate exactly once')).toBeLessThan(
      candidate.indexOf('Record immutable source identity before matrices'),
    );
    expect(candidate.indexOf('npm pack --ignore-scripts --pack-destination')).toBeGreaterThan(
      candidate.indexOf('ONLYOFFICE_SKIP_RELEASE_BUILD=1 pnpm build'),
    );
    expect(candidate).toContain('test -s dist/npm/public-api.js');
    expect(candidate).toContain('test -s dist/npm/public-api.d.ts');
    expect(candidate).toContain('npm install --global npm@11.17.0');
    expect(candidate).toContain('npm pack --ignore-scripts --pack-destination');
    expect(candidate.indexOf('Record immutable source identity before matrices')).toBeLessThan(
      candidate.indexOf('Run mandatory full-v5 Cloudflare and Broker matrix'),
    );
  });
  it('promotes only a successful named candidate with no runtime rebuild', () => {
    expect(promotion).toContain('actions: read');
    expect(promotion).toContain('fetch-depth: 0');
    expect(promotion).toContain('git merge-base --is-ancestor "${{ inputs.candidate_commit }}" origin/main');
    expect(promotion).not.toContain('test "${GITHUB_SHA}" = "${{ inputs.candidate_commit }}"');
    expect(promotion).toContain('candidate_run_id');
    expect(promotion).toContain('staging_run_id');
    expect(promotion).toContain('onlyoffice-runtime-stage-${{ inputs.candidate_commit }}-${{ inputs.staging_run_id }}');
    expect(promotion).toContain('staging evidence is not bound to this exact candidate run and runtime');
    expect(promotion).toContain('piwork_integration_run_id');
    expect(promotion).toContain('piwork_integration_run_attempt');
    expect(promotion).toContain("conclusion!=='success'");
    expect(promotion).toContain("['push','workflow_dispatch'].includes(r.event)");
    expect(promotion).toContain("r.head_branch!=='main'");
    expect(promotion).toContain('runtime-asset-version.mjs');
    expect(promotion).toContain('--no-traverse');
    expect(promotion).toContain('--transfers 16');
    expect(promotion).toContain('--checkers 32');
    expect(promotion).toContain("--exclude 'channels/**'");
    expect(promotion).toContain('--concurrency 16');
    expect(promotion).not.toMatch(/rclone copy candidate-input\/dist r2:onlyoffice-getpi-work/);
    expect(promotion).not.toContain('wrangler@4.114.0 deploy ');
    expect(promotion).toContain('wrangler@4.114.0 versions upload');
    expect(promotion).toContain('WRANGLER_OUTPUT_FILE_PATH');
    expect(promotion).toContain('"${PREVIOUS_WORKER_VERSION_ID}@100%"');
    expect(promotion).toContain('"${CANDIDATE_WORKER_VERSION_ID}@0%"');
    expect(promotion).toContain('"${CANDIDATE_WORKER_VERSION_ID}@100%"');
    expect(promotion).toContain('--worker-version-id "${CANDIDATE_WORKER_VERSION_ID}"');
    expect(promotion).toContain('--expected-worker-version-id "${CANDIDATE_WORKER_VERSION_ID}"');
    expect(promotion).toContain('Verify candidate Worker with release-pinned paths through a version override');
    expect(promotion).toContain(
      'https://onlyoffice.getpi.work/r/${ONLYOFFICE_SOURCE_RELEASE_ID}/onlyoffice-runtime-assets.json?override-readiness=',
    );
    expect(
      promotion.slice(
        promotion.indexOf('Verify candidate Worker with release-pinned paths through a version override'),
        promotion.indexOf('Activate only the fully compatible candidate Worker'),
      ),
    ).not.toContain('https://onlyoffice.getpi.work/onlyoffice-runtime-assets.json');
    expect(promotion).toContain('The unversioned root is intentionally still bound to the previous');
    const beforePointerCommit = promotion.slice(
      0,
      promotion.indexOf('Promote stable-v5 only after both Worker compatibility directions pass'),
    );
    expect(beforePointerCommit).toContain(
      'https://onlyoffice.getpi.work/r/${ONLYOFFICE_SOURCE_RELEASE_ID}/onlyoffice-runtime-assets.json?worker-readiness=',
    );
    expect(beforePointerCommit).not.toContain(
      'https://onlyoffice.getpi.work/onlyoffice-runtime-assets.json?worker-readiness=',
    );
    expect(beforePointerCommit).not.toContain('--verify-stable-root');
    expect(
      promotion.slice(promotion.indexOf('Promote stable-v5 only after both Worker compatibility directions pass')),
    ).toContain('--verify-stable-root');
    expect(promotion).toContain('stable-v5-release-cas');
    expect(promotion).toContain('candidate envelope, manifest source, or source package integrity mismatch');
    expect(promotion).toContain('manifest.sourceCommit!==process.env.CANDIDATE');
    expect(promotion).toContain('validateReleaseEnvelope');
    expect(promotion).toContain('bindExistingNpmPublication');
    expect(promotion).toContain('candidate npm registry/SLSA/public API evidence is stale or forged');
    expect(promotion).toContain('candidate npm package must be published with verified provenance before promotion');
    expect(promotion).toContain('new npm publication is not the exact candidate commit and source package');
    expect(promotion).toContain('staging evidence is not bound to the current immutable npm publication');
    expect(promotion).toContain('descriptor.npmPackage?.integrity!==npmRegistry.integrity');
    expect(promotion).toContain('Piwork integration descriptor is not bound to this runtime candidate');
    expect(promotion).toContain("run.event!=='workflow_dispatch'");
    expect(promotion).toContain('commits/${PIWORK_COMMIT}/pulls');
    expect(promotion).toContain('exactly one open same-repository PR on main');
    expect(promotion).toContain('.github/actions/setup-toolchain/action.yml');
    expect(promotion).toContain('scripts/verify-onlyoffice-release.mjs Makefile');
    expect(promotion).toContain('Piwork integration may not modify');
    expect(promotion).toContain('descriptor.runtimeIdentity?.sourceCommit!==envelope.source.gitCommit');
    expect(promotion).toContain('descriptor.runtimeIdentity?.hostBuildId!==envelope.runtime.protocolHostBuildId');
    expect(promotion).toContain('descriptor.releaseManifest?.hostBuildId!==envelope.runtime.hostBuildId');
    expect(promotion).toContain('candidate matrix evidence mismatch');
    expect(promotion).toContain('Require successful production staging evidence for this candidate');
    expect(promotion).toContain("r.path.endsWith('/stage-r2.yml')");
    expect(promotion).toContain("run.event!=='workflow_dispatch'");
    expect(promotion).toContain('/actions/runs/${PIWORK_RUN_ID}/attempts/${PIWORK_RUN_ATTEMPT}');
    expect(promotion).toContain('/actions/runs/${PIWORK_RUN_ID}/attempts/${PIWORK_RUN_ATTEMPT}/jobs?per_page=100');
    expect(promotion).toContain('run.run_attempt!==Number(process.env.PIWORK_RUN_ATTEMPT)');
    expect(promotion).toContain("deepVerifyRunAttempt:id('PIWORK_RUN_ATTEMPT')");
    expect(promotion).toContain('actions/create-github-app-token@fee1f7d63c2ff003460e3d139729b119787bc349');
    expect(promotion).toContain('ONLYOFFICE_RELEASE_READ_APP_ID');
    expect(promotion).toContain('ONLYOFFICE_RELEASE_READ_APP_PRIVATE_KEY');
    expect(promotion).toContain("if: steps.piwork-app.outputs.enabled == 'true'");
    expect(promotion).toContain('Configure both optional GitHub App secrets or neither');
    expect(promotion).toContain('PIWORK_READ_TOKEN: ${{ steps.piwork-read-token.outputs.token }}');
    expect(promotion).toContain('if test -n "${PIWORK_READ_TOKEN}"');
    expect(promotion).toContain('-H "Authorization: Bearer ${PIWORK_READ_TOKEN}"');
    expect(promotion).toContain('raw_get()');
    expect(promotion).toContain(
      'local args=(--fail --silent --show-error --retry 4 --retry-all-errors --connect-timeout 10 --max-time 30)',
    );
    expect(promotion).not.toContain('${RAW_HEADERS[@]}');
    expect(promotion).toContain("-H 'X-GitHub-Api-Version: 2022-11-28'");
    expect(promotion).toContain('x-ratelimit-limit=${limit:-unknown}');
    expect(promotion).toContain('Configure the optional Piwork read GitHub App');
    expect(promotion.match(/https:\/\/api\.github\.com\/repos\/agentbridges-ai\/pi-work/g)).toHaveLength(3);
    expect(promotion).toContain('OnlyOffice candidate integration / ${envelope.runtime.releaseId}');
    expect(promotion).toContain("descriptor.lifecycle!=='candidate'");
    expect(promotion).toContain('timeout-minutes: 180');
    expect(promotion).toContain('Persist or recover immutable pointer and Worker promotion intent');
    expect(promotion).toContain('promotion-intents/');
    expect(promotion).toContain('rclone cat "r2:onlyoffice-getpi-work/${INTENT_PATH}"');
    expect(promotion).toContain('if ! timeout 120 rclone cat');
    expect(promotion).toContain('return 1');
    expect(promotion).toContain('test -s "${temporary}"');
    expect(promotion).toContain('mv "${temporary}" promotion-intent.json');
    expect(promotion).toContain('stable-v5 is neither the recorded predecessor nor this candidate');
    expect(promotion).toContain('stable-v5 changed during production deployment; refusing activation');
    expect(promotion).toContain('Record immutable promotion receipt after stable-v5 convergence');
    expect(promotion.indexOf('Record immutable promotion receipt after stable-v5 convergence')).toBeGreaterThan(
      promotion.indexOf('Promote stable-v5 only after both Worker compatibility directions pass'),
    );
    expect(promotion).toContain('previousStable');
    expect(promotion).toContain("require('./promotion-intent.json').previousStable");
    expect(promotion).toContain('public promotion receipt bytes, digest, or immutable cache policy mismatch');
    expect(promotion).toContain('Compensate an incomplete pointer and Worker transaction');
    expect(promotion).toContain("['previous/previous','previous/split','previous/candidate','candidate/candidate']");
    expect(promotion).toContain(
      'const matchesPrevious=same(current,intent.previousStable),matchesCandidate=same(current,intent.runtime)',
    );
    expect(promotion).toContain(
      "const pointerState=matchesCandidate&&workerState==='candidate'?'candidate':matchesPrevious?'previous':matchesCandidate?'candidate':null",
    );
    expect(promotion.indexOf('const workerState=')).toBeLessThan(promotion.indexOf('const pointerState='));
    expect(promotion).toContain("fs.writeFileSync('.promotion-intent-adopted'");
    expect(promotion).toContain("fs.writeFileSync('.promotion-already-converged'");
    expect(promotion).toContain('No fully validated promotion intent was adopted by this run');
    expect(promotion).toContain('Production was already converged when this run adopted the intent');
    expect(promotion.indexOf("fs.writeFileSync('.promotion-intent-adopted'")).toBeGreaterThan(
      promotion.indexOf('immutable promotion intent ${field} mismatch'),
    );
    expect(promotion).toContain('previousVersionId');
    expect(promotion).toContain('candidateVersionId');
    expect(promotion).toContain('finalDeploymentId');
    expect(promotion).toContain('Retain production promotion receipt');
    expect(promotion).toContain('trustRoot');
    expect(promotion).toContain("channel:'stable-v5'");
    expect(promotion).toContain("staging:{runId:id('STAGING_RUN_ID')}");
    expect(promotion).toContain('must be a positive safe integer');
    expect(promotion).toContain('path:`promotions/');
    expect(promotion).toContain('promotion-receipt/metadata.json');
    expect(promotion).toContain('PROMOTION_RECEIPT_PATH="$(node -p');
    expect(promotion).toContain('export PROMOTION_RECEIPT_PATH PROMOTION_RECEIPT_FILE PROMOTION_RECEIPT_SHA256');
    expect(promotion).toContain('legacy-before.json');
    expect(promotion).toContain('Cache-Control: no-store');
    expect(promotion).toContain('stable-v5 did not converge to the exact no-store pointer');
    expect(promotion).toContain('Worker traffic and the unversioned root can converge at different');
    expect(promotion).toContain('promotion-worker=${CANDIDATE}-${GITHUB_RUN_ID}-${GITHUB_RUN_ATTEMPT}-${attempt}');
    expect(promotion).toContain('test "${attempt}" = 15');
    expect(promotion).not.toContain('--verify-pointers');
    expect(promotion.indexOf('verify-release-http')).toBeLessThan(
      promotion.indexOf('Promote stable-v5 only after both Worker compatibility directions pass'),
    );
    expect(promotion).not.toMatch(/pnpm build|release:build|hydrate-cloudflare|run-cloudflare|ONLYOFFICE_CF_MATRIX/);
  });
  it('rolls back compare-before-write and only stable-v5 with readback', () => {
    expect(rollback).toContain('EXPECTED_CURRENT_RELEASE_ID');
    expect(rollback).toContain('EXPECTED_CURRENT_MANIFEST_SHA256');
    expect(rollback).toContain('EXPECTED_MANIFEST_SHA256');
    expect(rollback).toContain('expected_current_manifest_sha256');
    expect(rollback).toContain('expected_manifest_sha256');
    expect(rollback).toContain('timeout-minutes: 180');
    expect(rollback).toContain('test "${GITHUB_REF}" = refs/heads/main');
    expect(rollback).toContain('target-promotion-receipts');
    expect(rollback).toContain('!positive(receipt.piwork?.deepVerifyRunAttempt)');
    expect(rollback).toContain('/actions/runs/${production_run_id}/attempts/${production_run_attempt}');
    expect(rollback).toContain("run.path==='.github/workflows/deploy-r2.yml'");
    expect(rollback).toContain(
      'rollback target was never activated by a successful exact protected production run attempt',
    );
    expect(rollback).toContain('test "${MANIFEST_SHA256}" = "${EXPECTED_MANIFEST_SHA256}"');
    expect(rollback).toContain('channels/stable-v5.json');
    expect(rollback).toContain('stable-v5-readback.json');
    expect(rollback).toContain('stable-v5-before-write.json');
    expect(rollback).toContain('stable-v5 changed during rollback verification; refusing activation');
    expect(rollback).toContain('retry_rclone()');
    expect(rollback).toContain('--v5-manifest manifest.json');
    expect(rollback).toContain('--remote-verification-mode full');
    expect(rollback).toContain('--concurrency 16');
    expect(rollback).toContain('--expected-worker-version-id "${CURRENT_WORKER_VERSION_ID}"');
    expect(rollback).toContain('--verify-stable-root');
    expect(rollback).toContain('Restore the original pointer if rollback verification fails');
    expect(rollback).toContain('Worker deployment changed during rollback verification; refusing activation');
    expect(rollback).toContain('CURRENT_WORKER_DEPLOYMENT_ID');
    expect(rollback).toContain('worker-after-rollback.json');
    expect(rollback).toContain('touch .rollback-committed');
    expect(rollback).toContain('rollbacks/${target.releaseId}');
    expect(rollback).toContain("authorization:{kind:'promotion-receipt'");
    expect(rollback).toContain('currentDeploymentId:process.env.CURRENT_WORKER_DEPLOYMENT_ID');
    expect(rollback.indexOf('touch .rollback-committed')).toBeLessThan(
      rollback.indexOf("fs.writeFileSync('rollback-receipt.json'"),
    );
    expect(rollback).toContain('Cache-Control: no-store');
    expect(rollback).toContain('stable-v5 rollback did not propagate to its immutable manifest');
    expect(rollback).toContain('cmp -s stable-v5-failed-rollback.json stable-v5.json');
    expect(rollback).toContain('cmp -s stable-v5-failed-rollback.json stable-v5-before.json');
    expect(rollback).not.toContain('CURRENT_RELEASE_ID=');
    expect(rollback).not.toContain('r2:onlyoffice-getpi-work/channels/stable.json');
  });
  it('audits the complete stable storage set without a mutable write', () => {
    expect(audit).toContain('schedule:');
    expect(audit).toContain('--v5-manifest manifest-public.json');
    expect(audit).toContain('--remote-verification-mode full');
    expect(audit).toContain('--concurrency 16');
    expect(audit).toContain('R2_AUDIT_READONLY_ACCESS_KEY_ID');
    expect(audit).toContain('degraded public-only mode');
    expect(audit).toContain('x-onlyoffice-worker-version');
    expect(audit).toContain('unversioned runtime root is not bound to stable-v5 CAS');
    expect(audit).toContain('promotion-receipt-names.txt');
    expect(audit).toContain('rollback-receipt-names.txt');
    expect(audit).toContain('promotion-evidence.json');
    expect(audit).toContain('rollback-evidence.json');
    expect(audit).toContain('!positive(receipt.piwork?.deepVerifyRunAttempt)');
    expect(audit).toContain('/actions/runs/${receipt_run_id}/attempts/${receipt_run_attempt}');
    expect(audit).toContain("'.github/workflows/deploy-r2.yml'");
    expect(audit).toContain("'.github/workflows/rollback-r2.yml'");
    expect(audit).toContain('Release/Worker pair authorization: passed');
    expect(audit).toContain('public release-pair receipt bytes, digest, immutable cache policy');
    expect(audit).not.toMatch(/rclone (?:copy|copyto) [^\n]+ r2:onlyoffice-getpi-work/);
  });
  it('pins every third-party action to a full commit SHA', () => {
    for (const workflow of [candidate, promotion, rollback, audit, npmPublication, pullRequestChecks])
      expect(workflow).not.toMatch(/uses:\s+[^\s]+@v\d/);
  });

  it('has an explicit current-build matrix mode that cannot rebuild the candidate', () => {
    expect(matrixRunner).toContain("ONLYOFFICE_CF_MATRIX_USE_CURRENT_BUILD === '1'");
    const currentBuildBranch = matrixRunner.slice(
      matrixRunner.indexOf('if (useCurrentBuild)'),
      matrixRunner.indexOf('} else if (!reusedPersistDirectory || refreshReusedState)'),
    );
    expect(currentBuildBranch).toContain('Current-build matrix input is missing');
    expect(currentBuildBranch).not.toContain("await run('pnpm', ['build'])");
    expect(currentBuildBranch).not.toContain('hydrate-cloudflare-matrix-fonts.mjs');
    expect(currentBuildBranch).not.toContain("await run('pnpm', ['release:build'])");
  });
});
