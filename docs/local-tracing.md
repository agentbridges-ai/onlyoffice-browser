# Local Cloudflare tracing

The Cloudflare runtime entrypoint for the OnlyOffice host is
`cloudflare/worker.ts`. Run it locally with Wrangler to get automatic local
invocation traces and Local Explorer:

```sh
pnpm run dev:trace
```

To use another local port, pass Wrangler arguments without an extra separator:

```sh
pnpm run dev:trace --port 8789
```

Open `/cdn-cgi/explorer` on the printed local URL, or use the
`/cdn-cgi/explorer/api` endpoint for local observability queries. Wrangler uses
local binding simulations by default; this command does not deploy, use the
remote R2 bucket, or move the Pi/Piwork runtime into Cloudflare.

The existing `test:e2e:cloudflare` matrix uses the same pinned Wrangler
version, so its local Worker invocations also have the tracing/Explorer
surface available while diagnosing a failing test.
