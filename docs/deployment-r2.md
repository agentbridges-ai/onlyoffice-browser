# Cloudflare R2 v5 release deployment

The production runtime is served by the `onlyoffice-browser-runtime` Worker:

- `https://onlyoffice.getpi.work` is the canonical resource, Broker, and
  installer origin.
- Twelve stable constellation hosts, from `https://aries.getpi.work` through
  `https://pisces.getpi.work`, provide isolated editor origins.
- The `onlyoffice-getpi-work` R2 bucket stores immutable release objects and the
  small mutable application shell.

Release Manifest v5 is the authoritative resource identity. It combines
file-level content-addressed storage with deterministic FastCDC representations
for assets at or above 8 MiB. This both enables local-change incremental updates
and keeps every Chrome Cache Storage write at or below 1 MiB. The all-in-one Office Pack remains a
v4 compatibility representation; it is not the canonical browser storage
format.

The canonical Broker accepts only objects declared by the pinned v5 manifest.
An object request is release-bound:

```text
/objects/<releaseId>/sha256/<digest>
```

The Worker resolves the digest to either a file/FastCDC blob or an Office Pack
segment only after checking that the requested release declares that exact
digest and byte length. Existing editors stay pinned to their release while a
new release is prepared and activated.

## R2 object layout

```text
blobs/sha256/<digest>
  Complete file objects and deterministic FastCDC chunks.

segments/sha256/<digest>
  Immutable Office Pack verification segments.

packages/sha256/<digest>.oobpack
  Complete v4 compatibility Office Pack.

releases/<releaseId>/manifest.json
  Release Manifest v5.

releases/<releaseId>/manifest-v4.json
  v4 compatibility manifest for the same release ID.

channels/stable-v5.json
  The only authoritative production activation pointer.

channels/stable.json
  Frozen v4 compatibility pointer for already-deployed legacy clients.
```

Every immutable filename contains its SHA-256 identity. Deployments are
additive and never run `rclone sync`, `--delete`, or `--delete-after`. Existing
release objects are not overwritten (`--ignore-existing`); an existing object
with incorrect bytes is detected by the mandatory remote SHA-256 gate.

The release workflow does not perform garbage collection. Retention and
mark-and-sweep must remain a separate operation that keeps all releases
referenced by Piwork, active editor leases, the current pointers, the newest
three releases, and the minimum retention window. It must also treat every
unfinalized `promotion-intents/**` record (both its candidate and predecessor),
every staged candidate still inside the staging-retention window, and every
valid retained `promotions/**` receipt as reachability roots. An intent is
finalized only when its digest-bound receipt exists; abandoning an intent or a
stage requires an explicit audited decision after its recovery window expires.

## v5 and v4 pointer contract

`channels/stable-v5.json` must point to
`/releases/<releaseId>/manifest.json`, and its `manifestSha256` must match the
exact manifest bytes. That manifest must have `version: 5`. New installations
and update checks use this pointer.

`channels/stable.json` must point to
`/releases/<releaseId>/manifest-v4.json`, and its digest must match the exact
v4 manifest bytes. It is a read-only compatibility anchor: normal v5 releases
validate it before and after activation but never rewrite it. Its release may
therefore be older than `stable-v5.json`; the referenced v4 release and all of
its immutable objects remain retained so legacy clients continue to work.

`stable-v5.json` is the single activation write. R2 strong consistency makes
that one `PutObject` an atomic replacement. The workflow no longer tries to
advance `stable.json` and `stable-v5.json` sequentially, because two individually
atomic writes do not form a cross-object transaction and could leave protocols
on an accidental half-switched release. Pointer responses use
`Cache-Control: no-store`.

An intentional legacy-pointer migration is a separate, audited compatibility
operation. It must validate the target v4 manifest and retained object set and
must not be combined with a normal v5 release activation.

## Source identity and mandatory local matrix

Every runtime candidate is bound to one reproducible source identity before any
R2 upload:

- the exact checked-out Git commit, which must equal `GITHUB_SHA`;
- the same 40-character commit in the immutable v5 manifest's `sourceCommit`;
  it participates in the release identity, so changing the candidate commit
  naturally changes `releaseId` and the manifest digest;
- the complete SHA-256 of `pnpm-lock.yaml`;
- the complete SHA-256 of `dist/onlyoffice-runtime-assets.json`;
- the v5 release ID and complete release-manifest SHA-256;
- the exact source-package tgz filename and SHA-512 SRI produced by the
  candidate commit. This is build evidence, not the npm registry identity.
- SHA-256 digests of `dist/npm/public-api.js` and `public-api.d.ts`, plus an
  npm-registry snapshot when that version already exists. The snapshot binds
  registry integrity, `gitHead`, and the decoded npm SLSA identity. Reusing an
  existing npm version is permitted only when both public API files are
  byte-equivalent; the source tarball SRI is deliberately not compared.

The hand-off deliberately distinguishes two Host identities. The v5 manifest
and release envelope call the SHA-256 aggregate of the Host files
`hostBuildId`. They call the Host protocol identity reported to an embedding
client `protocolHostBuildId` (for example `office-host-0.5.12-r1`). Piwork maps
the latter to `runtimeIdentity.hostBuildId` and the former to
`releaseManifest.hostBuildId`; swapping them makes the integration invalid.

`candidate-r2.yml` builds and hydrates the runtime exactly once. Both mandatory
Cloudflare/Broker matrices use
`ONLYOFFICE_CF_MATRIX_USE_CURRENT_BUILD=1`; that mode validates the retained
`dist` and release tree and is forbidden from rebuilding them. The first pass
runs the complete full-v5 matrix and seeds retained local state. The second pass
reuses those exact bytes for the fault-injected segment-stream test. After each
matrix, the workflow recomputes the source identity and fails if any byte-level
identity changed.

The release-manifest CLI accepts `--source-commit` (or
`ONLYOFFICE_SOURCE_COMMIT`). CI always passes the candidate `GITHUB_SHA`
explicitly. A local call may omit it only when the checked-out `HEAD` is a full
40-character commit and the tracked worktree and index are unchanged; otherwise
the builder fails closed instead of assigning uncommitted bytes to a commit.

On success, GitHub Actions retains
`onlyoffice-runtime-candidate-<gitCommit>` for 30 days. The artifact contains
the exact `dist`, `.onlyoffice-release`, source-package tgz, release envelope,
`source-identity.json`, `full-v5-cloudflare-broker-matrix.json`, and
`fault-injected-proxy-matrix.json`. A successful process exit without this
complete artifact is not a promotion input.

## FastCDC bounded-write policy

Every asset at or above 8 MiB uses this fixed policy:

```text
fastcdc-v2020
min 64 KiB / average 256 KiB / max 1 MiB
normalization 1 / seed 0
```

Production explicitly sets:

```text
ONLYOFFICE_FASTCDC_EVIDENCE_MODE=automatic
```

The verifier requires every >=8 MiB asset to carry complete, contiguous FastCDC
metadata and rejects FastCDC on smaller assets. Historical release evidence is
still collected as an incremental-efficiency benchmark, but it cannot disable
the bounded-write safety representation.

## Candidate and production order

The release train deliberately separates candidate creation from activation.

`candidate-r2.yml` performs these steps on `main`, either from a push or a
manual dispatch:

1. Derive `SOURCE_DATE_EPOCH` from the checked-out commit, install dependencies,
   verify the signed, immutable, attested x2t upstream release, run tests and
   type checking, build the Host/runtime, and hydrate the pinned immutable font
   release. A cache miss performs a read-only download from the fixed public
   release and verifies the pinned font manifest plus every object digest
   before repopulating the immutable cache key.
2. Build deterministic v5/v4 manifests, pointers, blobs, chunks, segments and
   package, passing `--source-commit "${GITHUB_SHA}"` explicitly; create the
   source-package tgz and one cross-surface release envelope.
3. Verify all local identities and run the full and fault-injected matrices
   against those exact retained bytes.
4. Retain the promotion artifact last. The candidate receives no R2 or Worker
   credentials and performs no production write. Its only production HTTP
   dependency is the digest-pinned, read-only font fallback on a cache miss.

`stage-r2.yml` is a manual, production-environment staging transaction. It
requires `STAGE`, the same candidate commit, and the successful candidate run.
It repeats the envelope, npm registry/SLSA, source-package, and matrix checks
without rebuilding, then uploads only `.onlyoffice-release/{blobs,releases,packages,segments}`
with `--ignore-existing` and fully re-hashes the remote CAS. It does not write a
channel, copy `dist` compatibility-root files, or deploy a Worker. It verifies
the release Host, Broker, and release-bound object routes through the _current_
production Worker and retains an evidence artifact. This proves current-Worker
compatibility for Piwork online candidate verification; it does not validate a
candidate Worker change, which is verified only during promotion before stable
activation.

`deploy-r2.yml` is a manual, production-environment approval transaction:

1. Require an explicit `PROMOTE`, a 40-character candidate commit, and the run
   ID of a successful `candidate-r2.yml` run for that commit on `main`.
2. Require a successful `stage-r2.yml` run on `main` for the same candidate
   SHA. Promotion re-verifies the remote CAS rather than treating staging as a
   substitute for its production check.
3. Require a successful, explicitly dispatched Piwork `deep-verify` run for an
   exact commit and the successful job named `OnlyOffice candidate integration /
<releaseId>`. The descriptor lifecycle must be `candidate` and binds
   `runtimeIdentity.sourceCommit` to the candidate commit. That commit must have
   one open same-repository PR to `main` whose head still matches; the promotion
   gate byte-compares Piwork's verifier and deep-verify workflow with `main`, so
   the integration PR cannot weaken its own evidence.
4. Download the named artifact, check out the same commit, and verify the
   envelope, lockfile, runtime/release digests, source-package tgz SRI, and both
   matrix evidence files. It never rebuilds the runtime.
5. Snapshot the current pointer and 100%-active Worker version. Upload the
   source-matched candidate as an undeployed Worker version, then persist an
   immutable promotion intent binding the predecessor/candidate pointer and
   both Worker version IDs before changing traffic.
6. Upload only the immutable CAS from the retained artifact and fully hash it
   remotely. Unversioned runtime-root requests resolve through `stable-v5` to
   release CAS; promotion never overwrites the legacy mutable bucket root.
7. Place a new candidate Worker at 0% beside the predecessor at 100%, wait for
   version-override propagation, and verify the candidate Worker with both the
   candidate and predecessor Host, Broker, Range, and stable-root routes. The
   response version-metadata header proves every override reached the intended
   Worker. A recovered candidate already at 100% is verified in place rather
   than temporarily downgraded.
8. Move the verified candidate Worker to 100%, verify normal traffic still
   serves both immutable releases, then compare-before-write and replace only
   `stable-v5.json`. Require its public `no-store` bytes, immutable manifest,
   and unversioned runtime-root digest to converge while `stable.json` remains
   byte-for-byte frozen. Any pre-commit failure compensates the pointer first
   and explicitly rolls back to the recorded predecessor Worker version.
9. Only after pointer and Worker convergence, write a content-addressed promotion receipt
   under `promotions/<releaseId>/`, verify its public immutable HTTPS bytes and
   SHA-256, and retain a small GitHub artifact. Its URL and digest are the
   inputs a later Piwork supported descriptor records. The protected production
   receipt records `channel: stable-v5`, the successful staging run, both
   Worker version IDs, the final deployment ID, and the stable-root CAS mode. The
   protected production workflow and immutable R2 CAS are the trust root; it is not an independent
   GitHub artifact/Sigstore attestation.

The protected `production` environment needs:

- `R2_ACCESS_KEY_ID`
- `R2_SECRET_ACCESS_KEY`
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`
- optional `ONLYOFFICE_RELEASE_READ_APP_ID`
- optional `ONLYOFFICE_RELEASE_READ_APP_PRIVATE_KEY`

The last two recommended secrets must be configured together and identify a
GitHub App installed for `agentbridges-ai/pi-work`. The promotion workflow mints
a short-lived token restricted to that repository; the installation needs only
Actions: read, Contents: read, and Pull requests: read (plus GitHub's implicit
Metadata access). Do not replace it with a PAT. When neither secret is present,
promotion remains operational and makes only three anonymous public REST calls:
the exact run attempt, that attempt's jobs, and the commit's pull requests. A
failed or rate-limited request reports `x-ratelimit-limit`,
`x-ratelimit-remaining`, and `x-ratelimit-reset` and recommends configuring the
App. Promotion binds Piwork evidence to the exact deep-verify run ID and run
attempt, including that attempt's jobs endpoint.

Candidate creation requires no GitHub environment secrets. Production approval
policy is an organization decision; do not hard-code a self-review restriction
for this single-administrator repository.

## rclone 501 behavior

Cloudflare's S3-compatible endpoint can return a transient `501 Not
Implemented` for an individual transfer while a later rclone attempt succeeds.
The workflows use command timeouts and at most four bounded rclone attempts.

A zero rclone exit status after retry means only that the copy command
finished. It is never treated as release verification. Worker deployment stays
blocked until the subsequent independent `rclone cat` pass has re-read and
SHA-256-verified the complete v5 storage set. This also catches an incorrect
pre-existing object skipped by `--ignore-existing`.

## Worker verification contract

Before activation, CI proves:

- `/r/<releaseId>/office-host.html` returns `200` without a redirect, the exact
  release identity, immutable caching, the HTML MIME type, and isolated-origin
  headers.
- `/r/<releaseId>/resource-broker.html` has a restrictive CSP with no wildcard
  source and explicitly allows Piwork, the canonical origin, and the twelve
  controlled editor origins as frame ancestors.
- `/objects/<releaseId>/sha256/<digest>` returns the manifest-declared byte
  length, `X-Content-SHA256`, immutable caching, CORS/CORP headers, correct
  `HEAD`, correct single-range `206`, and correct unsatisfiable-range `416`.
- The returned range bytes exactly match the locally verified immutable object.
- `stable-v5.json` is `no-store`, identifies the expected new release, and
  leads to a v5 manifest whose exact bytes match its declared SHA-256.
- The frozen `stable.json` response is `no-store`, stays byte-for-byte
  unchanged across activation, and continues to resolve to a digest-valid v4
  compatibility manifest. It is not required to share the v5 release ID.

The direct R2 pass and the HTTP pass are intentionally separate: the first
proves stored bytes; the second proves Worker routing and HTTP semantics.

## Recovery deployment

Recovery uses the same retained identities; it does not introduce an alternate
deployment path:

- If a candidate artifact is still retained, rerun `deploy-r2.yml` with its
  original commit and run ID. The workflow re-verifies every identity before
  any mutable write.
- If the artifact expired, rerun `candidate-r2.yml` for the exact `main` commit,
  then promote the new successful run. Do not reconstruct a promotion artifact
  by hand.
- To restore a previous runtime, run `rollback-r2.yml` from `main` with the
  target immutable `releaseId`, its original promotion receipt's exact manifest
  SHA-256, the observed current `stable-v5` release ID and manifest SHA-256,
  and `ROLLBACK`.
  Compare-before-write prevents a blind overwrite; the workflow verifies the
  target v5 manifest and complete R2 set, requires a digest-bound promotion
  receipt backed by its exact successful protected production run attempt, then
  preflights Host/Broker/Range compatibility against the exact current Worker
  before the pointer write. Post-write it verifies the stable-root CAS bytes
  and unchanged Worker deployment, then writes an immutable rollback receipt
  binding the from/target pointers, promotion authorization, Worker deployment,
  and rollback run. Promotion is the sole trust root; rollback receipts only
  form a verifiable chain to it.
- A same-target rollback rerun is receipt reconciliation: it re-verifies the
  already-converged pointer and emits the missing immutable receipt. Once the
  target pointer and unchanged Worker have been verified, a later receipt upload
  or public-readback failure does not compensate a legitimate rollback; rerun
  with the exact now-current pointer identity to reconcile the receipt.
- To retry a transient npm release failure before publication completes,
  dispatch `release-npm.yml` from the same existing signed tag ref. Existing npm
  versions remain verification-only and are never overwritten.

An emergency manual activation requires an incident record and an independent
reviewer. It must reproduce the candidate envelope and both matrix evidence
files, use bounded additive R2 copies, fully re-hash the remote storage set,
verify the exact Host route before activation, snapshot the frozen legacy
pointer, and perform only the single `stable-v5.json` write. Never use a
destructive sync command or write `stable.json` during a normal v5 recovery.

Configure the `r2` remote as Cloudflare S3 with a bucket-scoped Object Read &
Write token, the account R2 endpoint, `acl = private`, and
`no_check_bucket = true`. R2 is required because x2t WebAssembly and the Office
resource set exceed Workers Static Assets' per-object constraints.

## Cache headers

- `channels/*.json`: `no-store`.
- `blobs/sha256/*`, `segments/sha256/*`,
  `packages/sha256/*`, release manifests, `/objects/<releaseId>/...`, and
  `/r/<releaseId>/...`: one year plus `immutable`.
- Unversioned application-shell documents: revalidate.
- Office resource responses: `no-transform`, with exact MIME, length, Range,
  CORS, CORP, and content identity headers.

An editor instance never follows `stable-v5.json` after it starts. It continues
to request its pinned release ID, so publishing v5 cannot mix bytes into an
already running v4/v5 editor.

## Release train

Runtime and npm have separate immutable identities. A runtime candidate is a v5
CAS `releaseId`; it may advance for Host-only changes while retaining an already
published npm version. An npm release is an exact tgz SRI plus source `gitHead`.

`candidate-r2.yml` builds once, runs all runtime gates, and retains an artifact
without production credentials or writes; a font-cache miss may read only the
fixed public immutable font release. `stage-r2.yml` makes
only its immutable CAS available to the current production Worker and emits
compatibility evidence, allowing Piwork to online-verify the candidate without
an activation. `deploy-r2.yml` requires that stage run, repeats full remote
verification, transactionally verifies the candidate Worker at 0%, activates
it at 100%, then moves only `stable-v5`. Runtime-root aliases follow release
CAS instead of mutable compatibility-root objects. `rollback-r2.yml` binds the
target to its recorded manifest digest and successful protected promotion,
preflights the current Worker, restores only that pointer, and records a
promotion-chained immutable rollback receipt.

## Required GitHub configuration

Repository settings are part of the release control plane and cannot be
enforced by this local change alone:

- Protect `main` with pull requests, verified commit signatures, the `verify`
  job from `Pull request checks`, conversation resolution, and no direct-push
  bypass for routine maintainers.
- Protect `v*` tags against deletion or movement and require verified signed
  annotated tags. The npm workflow rejects lightweight, unverified, mismatched,
  and wrong-commit tags even if a ruleset is misconfigured.
- Do not configure a candidate environment or candidate R2 credentials: it is
  artifact-only. The production workflow is the first writer to production R2.
- Without audit secrets, `audit-r2.yml` performs a degraded public HTTP check.
  Configure the optional audit-only read credentials to enable its complete R2
  re-hash; neither mode has a channel or Worker mutation step.
- Give the `production` environment the promotion/rollback R2 credentials and
  Worker deployment token. Allow only `main` and protected `v*` tags: staging,
  promotion, and rollback use `main`, while npm trusted publication runs
  from the exact signed tag. Use the review policy that fits the repository's
  available administrators.
- Configure npm Trusted Publishing for package
  `@agentbridges-ai/onlyoffice-browser`, repository
  `agentbridges-ai/onlyoffice-browser`, workflow
  `.github/workflows/release-npm.yml`, and environment `production`. Do not add
  a long-lived npm token.
- Keep every third-party Action pinned to a full commit SHA and review dependency
  updates as supply-chain changes.

`audit-r2.yml` provides the separate read-only full publication audit. It
authorizes the active release/Worker pair either directly through a successful
promotion receipt or through a successful rollback receipt whose target digest
and authorization chain resolve to that valid promotion receipt. CAS
garbage collection must remain an independently approved mark-and-sweep job and may run only after
enumerating `stable-v5`, the frozen legacy pointer, retained editor leases, the
newest three releases, the minimum retention window, and every release still
pinned by supported Piwork versions. Active promotion intents, their candidate
and predecessor releases, unexpired staged candidates, and digest-valid
promotion receipts and valid promotion-chained rollback receipts are additional
roots. GC must not infer that an intent is
abandoned merely because its originating workflow stopped; only a matching
receipt finalizes it, and abandonment is a separate recorded operation.
