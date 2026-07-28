# Cloudflare wildcard runtime deployment

The production runtime is served by the `onlyoffice-browser-runtime` Worker:

- `https://onlyoffice.getpi.work` is the canonical shared-asset origin.
- `https://office-editor-<session>.getpi.work/office-host.html` gives every
  editor its own host origin and renderer lifecycle.
- The `onlyoffice-getpi-work` R2 bucket stores the compact runtime and generated
  font overlay.

The Worker keeps the small origin-bound bootstrap documents, Vite host chunks,
and Worker scripts on each editor origin. All other OnlyOffice resources are
redirected to one build-versioned canonical URL. As a result, new editor origins
reuse the browser's existing HTTP cache instead of downloading another copy of
the SDK, dictionaries, fonts, and WebAssembly.

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
wildcard hostnames that do not start with `office-editor-`.

## Deployment

The GitHub workflow tests, builds, and compacts the runtime, synchronizes
`dist/` while preserving the separately generated font overlay and Piwork
plugin prefix, deploys the Worker with a content-derived revision of the shared
runtime files, and verifies the production version header and isolated host
route. Generated timestamps and changes limited to the demo UI do not
invalidate the large shared runtime cache.

Required production-environment secrets:

- `R2_ACCESS_KEY_ID`
- `R2_SECRET_ACCESS_KEY`
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`

Normal releases are performed by `.github/workflows/deploy-r2.yml` after a push
to `main` (or by `workflow_dispatch`). Local deployment is only a recovery path:

```bash
pnpm build
rclone sync dist r2:onlyoffice-getpi-work \
  --fast-list --checkers 32 --transfers 16 --delete-after \
  --exclude "/fonts/**" \
  --exclude "/onlyoffice-browser-font-assets.json" \
  --exclude "/onlyoffice-browser-font-source-map.json" \
  --exclude "/sdkjs/common/AllFonts.js" \
  --exclude "/sdkjs/common/Images/fonts_thumbnail*.png" \
  --exclude "/server/FileConverter/bin/AllFonts.js" \
  --exclude "/server/FileConverter/bin/font_selection.bin" \
  --exclude "/onlyoffice-plugin/**"
ASSET_VERSION="$(
  node scripts/runtime-asset-version.mjs \
    --root dist \
    --font-manifest .onlyoffice-font-assets/onlyoffice-browser-font-assets.json
)"
npx --yes wrangler@4.114.0 deploy --var "ASSET_VERSION:${ASSET_VERSION}"
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

Canonical URLs include `?__oobv=<manifest revision>` and are returned with a
one-year immutable cache policy. Unversioned canonical documents and per-editor
bootstrap files revalidate so a deployment cannot strand a new editor on stale
HTML. The generated fonts use the same canonical versioned URLs and Chrome's
native HTTP cache.

This is a weak-network design, not a strict offline PWA. A new editor origin
still needs a small request/redirect handshake and its origin-bound bootstrap
files, while large runtime resources are reused from the canonical HTTP cache.
