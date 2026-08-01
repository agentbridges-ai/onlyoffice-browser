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
three releases, and the minimum retention window.

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

Every production release is bound to a reproducible source identity before any
R2 upload:

- the exact checked-out Git commit, which must equal `GITHUB_SHA`;
- the complete SHA-256 of `pnpm-lock.yaml`;
- the complete SHA-256 of `dist/onlyoffice-runtime-assets.json`;
- the v5 release ID and complete release-manifest SHA-256.

The v5 manifest's `runtimeManifestSha256` and the
`onlyoffice-runtime-assets.json` asset digest must both match the built runtime
manifest. After the matrix rebuilds the product, CI recomputes all fields and
fails if any value differs.

Before any production pointer is changed, CI runs
`scripts/run-cloudflare-local-matrix.mjs` with:

```text
ONLYOFFICE_CF_MATRIX_FORCE_FULL=1
ONLYOFFICE_CF_MATRIX_GREP=
ONLYOFFICE_CF_MATRIX_REUSE_BUILD=0
ONLYOFFICE_CF_MATRIX_KEEP_STATE=0
```

The gate asserts these values before invoking the runner. A grep-filtered case,
the synthetic Broker fixture, a route-only run, or a reused build cannot satisfy
the production gate. The full-v5 matrix uses the production Worker and R2
bindings locally, performs a fresh deterministic build, installs the real
resource set once, clears HTTP cache while retaining canonical Cache Storage,
opens three isolated editor origins, exercises restart/offline Word,
Spreadsheet and Presentation flows, and validates incremental update,
release pinning, rollback, Range and CSP behavior.

On success, GitHub Actions retains two JSON files for 90 days in the
`full-v5-cloudflare-broker-matrix-<run>-<attempt>` artifact:

- `source-identity.json`;
- `full-v5-cloudflare-broker-matrix.json`.

The latter records the non-downgraded matrix policy, success result, workflow
run identity, and the post-matrix source/release identity. A successful process
exit without this artifact is not publishable evidence.

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

## Production deployment order

`.github/workflows/deploy-r2.yml` is the production authority. It performs the
following ordered transaction:

1. Derive `SOURCE_DATE_EPOCH` from the checked-out commit, install dependencies,
   run tests/type checking, build the shell and Host, and hydrate the pinned
   immutable font release.
2. Build deterministic `manifest.json` v5, `manifest-v4.json`, both pointer
   candidates, complete file blobs, bounded FastCDC chunks, package
   segments, and the complete compatibility package.
3. Run `scripts/verify-release-publication.mjs` locally. It checks the v5/v4
   release identity, both pointer digests, storage-set digest, package
   coverage, FastCDC evidence mode, and every local immutable object's complete
   SHA-256 and byte length.
4. Bind the build to Git commit, lockfile, runtime-manifest and release-manifest
   SHA-256 identities.
5. Run the mandatory, unfiltered, fresh-build full-v5 local Cloudflare/Broker
   matrix. Recompute the bound identities and retain the successful JSON
   evidence artifact.
6. Upload only missing immutable objects with
   `rclone copy --ignore-existing`. Upload the root compatibility shell without
   any delete operation and preserve the Piwork `/onlyoffice-plugin/` overlay.
7. Run `scripts/verify-release-publication.mjs` with the R2 remote
   `r2:onlyoffice-getpi-work`. This streams every declared whole-file blob,
   every declared FastCDC chunk, every package segment, the complete package,
   and both release manifests back from R2 with `rclone cat`, then recomputes
   full SHA-256 and byte length.
8. Deploy the Worker only after the complete R2 verification succeeds.
9. Before activation, run
   `scripts/verify-release-http.mjs`. It verifies the release-pinned Host,
   Broker CSP, release-bound object headers, and `200`/`206`/`416` Range
   semantics against the newly deployed Worker.
10. Read, validate and snapshot the existing `stable.json` v4 pointer and its
    referenced manifest. The HTTP bytes must match the direct R2 bytes.
11. Atomically write and read back only `stable-v5.json`.
12. Re-run the HTTP verifier, prove `stable-v5.json` exactly matches the new
    candidate twice, and prove `stable.json` remains byte-for-byte unchanged
    and still resolves to a digest-valid v4 manifest.

Required production-environment secrets:

- `R2_ACCESS_KEY_ID`
- `R2_SECRET_ACCESS_KEY`
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`

## rclone 501 behavior

Cloudflare's S3-compatible endpoint can return a transient `501 Not
Implemented` for an individual transfer while a later rclone attempt succeeds.
The workflow has bounded rclone retries plus three bounded outer attempts.

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

Re-running the production workflow with `workflow_dispatch` is the preferred
recovery path because it preserves the source-identity and matrix evidence.
A manual local deployment is an emergency path only. Authentication must be
checked first with `npx wrangler whoami`, and the same full-v5 matrix must pass
from a fresh build before any remote mutation:

```bash
ONLYOFFICE_CF_MATRIX_FORCE_FULL=1 \
ONLYOFFICE_CF_MATRIX_GREP= \
ONLYOFFICE_CF_MATRIX_REUSE_BUILD=0 \
ONLYOFFICE_CF_MATRIX_KEEP_STATE=0 \
node scripts/run-cloudflare-local-matrix.mjs
```

The operator must save equivalent source-identity and successful matrix JSON
evidence with the incident record. Starting from the matrix-verified,
font-hydrated `dist` directory:

```bash
pnpm release:build

node scripts/verify-release-publication.mjs \
  --release-root .onlyoffice-release \
  --expected-package-version 0.5.7 \
  --fastcdc-evidence-mode forbid

rclone copy .onlyoffice-release/blobs \
  r2:onlyoffice-getpi-work/blobs --ignore-existing
rclone copy .onlyoffice-release/releases \
  r2:onlyoffice-getpi-work/releases --ignore-existing
rclone copy .onlyoffice-release/packages \
  r2:onlyoffice-getpi-work/packages --ignore-existing
rclone copy .onlyoffice-release/segments \
  r2:onlyoffice-getpi-work/segments --ignore-existing
rclone copy dist r2:onlyoffice-getpi-work \
  --exclude "/onlyoffice-plugin/**"

node scripts/verify-release-publication.mjs \
  --release-root .onlyoffice-release \
  --expected-package-version 0.5.7 \
  --fastcdc-evidence-mode forbid \
  --remote r2:onlyoffice-getpi-work

ASSET_VERSION="$(
  node scripts/runtime-asset-version.mjs \
    --root dist \
    --font-manifest .onlyoffice-font-assets/onlyoffice-browser-font-assets.json
)"
npx wrangler deploy --var "ASSET_VERSION:${ASSET_VERSION}"

node scripts/verify-release-http.mjs \
  --release-root .onlyoffice-release \
  --expected-package-version 0.5.7 \
  --fastcdc-evidence-mode forbid

mkdir -p .onlyoffice-recovery
rclone cat r2:onlyoffice-getpi-work/channels/stable.json \
  > .onlyoffice-recovery/legacy-stable.json
# Validate the frozen pointer and its manifest exactly as the workflow does.

rclone copyto .onlyoffice-release/channels/stable-v5.json \
  r2:onlyoffice-getpi-work/channels/stable-v5.json
rclone cat r2:onlyoffice-getpi-work/channels/stable-v5.json |
  cmp --silent .onlyoffice-release/channels/stable-v5.json -
rclone cat r2:onlyoffice-getpi-work/channels/stable.json |
  cmp --silent .onlyoffice-recovery/legacy-stable.json -

node scripts/verify-release-http.mjs \
  --release-root .onlyoffice-release \
  --expected-package-version 0.5.7 \
  --fastcdc-evidence-mode forbid
```

Do not omit source-identity binding, successful full-v5 matrix evidence, remote
verification, pre-activation HTTP verification, or the legacy-pointer
unchanged check during recovery. Do not write `stable.json` as part of a normal
recovery release, and do not use a destructive sync command.

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
