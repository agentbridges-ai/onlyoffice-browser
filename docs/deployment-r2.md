# Cloudflare fixed-pool runtime deployment

The production runtime is served by the `onlyoffice-browser-runtime` Worker:

- `https://onlyoffice.getpi.work` is the canonical shared-asset origin.
- Twelve stable constellation hosts, from `https://aries.getpi.work` through
  `https://pisces.getpi.work`, provide one leased origin and renderer lifecycle
  per active editor.
- The `onlyoffice-getpi-work` R2 bucket stores the compact runtime and generated
  font overlay.

Release Manifest v4 combines the complete SDK, Word/Spreadsheet/Presentation
components, x2t runtime, dictionaries, and every shipped font into one
deterministic `office-resources.oobpack`. The Worker keeps the small
origin-bound bootstrap documents, Vite host chunks, and Worker scripts on each
editor origin. The installer transfers cacheable 24 MiB verification segments
from that immutable pack. The editor Service Worker resolves shared resource
URLs through those verified segments, so the product has one installation, one
progress value, and one integrity identity without changing the URLs expected
by OnlyOffice.

The content-addressed v3 blobs remain additive compatibility objects for older
clients and rollback. New v4 clients transfer one logical package in bounded
segments rather than issuing one installation request per resource.

R2 is required because the x2t WebAssembly binary exceeds the individual asset
limit of Workers Static Assets. The Worker binding also avoids depending on R2
custom-domain ownership state.

## Cloudflare resources

The checked-in `wrangler.jsonc` owns both zone routes and the R2 binding. The
zone must contain proxied DNS records for:

```text
onlyoffice.getpi.work
*.getpi.work
```

Both records may use the same placeholder origin because the Worker routes
intercept requests before an origin fetch. Editor hosts stay directly below
`getpi.work`, so Cloudflare Universal SSL covers them without Total TLS or an
Advanced Certificate Manager subscription. The Worker rejects first-level
wildcard hostnames outside the fixed constellation set. Legacy
`office-editor-<session>.getpi.work` hosts remain routable during rollback, but
new clients never allocate them.

## Deployment

The GitHub workflow tests, builds, and compacts the runtime, builds and hashes
the deterministic Office Pack, uploads missing immutable package/blob/release
objects, deploys the Worker, verifies package Range behavior and isolated-host
loading, and switches `channels/stable.json` last.

Required production-environment secrets:

- `R2_ACCESS_KEY_ID`
- `R2_SECRET_ACCESS_KEY`
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`

Normal releases are performed by `.github/workflows/deploy-r2.yml` after a push
to `main` (or by `workflow_dispatch`). Local deployment is only a recovery path:

```bash
pnpm build
pnpm release:build
rclone copy .onlyoffice-release/packages r2:onlyoffice-getpi-work/packages \
  --fast-list --checkers 16 --transfers 4 --ignore-existing
rclone copy .onlyoffice-release/releases r2:onlyoffice-getpi-work/releases \
  --fast-list --checkers 16 --transfers 8 --ignore-existing
rclone check .onlyoffice-release/packages r2:onlyoffice-getpi-work/packages \
  --one-way --checksum
ASSET_VERSION="$(
  node scripts/runtime-asset-version.mjs \
    --root dist \
    --font-manifest .onlyoffice-font-assets/onlyoffice-browser-font-assets.json
)"
npx --yes wrangler@4.114.0 deploy --var "ASSET_VERSION:${ASSET_VERSION}"
rclone copyto .onlyoffice-release/channels/stable.json \
  r2:onlyoffice-getpi-work/channels/stable.json
```

Configure the `r2` remote as Cloudflare S3 with a bucket-scoped Object Read &
Write token, the account R2 endpoint, `acl = private`, and
`no_check_bucket = true`. The command mirrors the compact runtime while
preserving overlay-owned paths. Generated fonts are a separately managed R2
overlay and are not present in a clean GitHub checkout; the runtime workflow
preserves them rather than pretending it can regenerate or upload them. Piwork
owns a separate `/onlyoffice-plugin/` deployment overlay, so runtime syncs
preserve that prefix as well.

## Cache model

The package URL is
`/p/<releaseId>/office-resources.oobpack`; `?segment=segment-NNN` exposes each
24 MiB package segment as a separate immutable HTTP cache entry, while the base
URL retains single-range compatibility. The package SHA-256 and every segment
SHA-256 are part of the release identity. Unversioned
canonical documents and fixed-pool bootstrap files revalidate so a deployment
cannot strand a new editor on stale HTML.

Every fixed editor origin stores each successfully SHA-256-verified package
segment in its own Cache Storage and reuses it across later editors assigned to
that origin. This is deliberately origin-owned instead of relying on an
unobservable cross-origin HTTP-cache warmup. Chrome does not reliably
synthesize an offline Range response from a very large cached object, so the
editor Service Worker reads the verified segment and slices the requested
resource itself. It retains only four segment buffers in memory. If package
discovery or a segment request fails, v3 immutable resource URLs remain the
compatibility fallback.
