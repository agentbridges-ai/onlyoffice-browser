# Browser Office Editor

[![CI](https://github.com/agentbridges-ai/onlyoffice-browser/actions/workflows/ci.yml/badge.svg)](https://github.com/agentbridges-ai/onlyoffice-browser/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@agentbridges-ai/onlyoffice-browser.svg)](https://www.npmjs.com/package/@agentbridges-ai/onlyoffice-browser)
[![license](https://img.shields.io/github/license/agentbridges-ai/onlyoffice-browser.svg)](LICENSE)

A browser-only Office preview/edit component powered by OnlyOffice 9.3 and `onlyoffice-x2t-wasm`.

The component opens DOCX, XLSX, PPTX, and CSV documents inside a host-provided DOM container. Conversion, editing, saving, and export happen in an independent-origin sandboxed editor host iframe. It does not require OnlyOffice DocumentServer, backend sessions, document uploads, or user accounts.

## What This Provides

- JS/TS API: `createOfficeEditor(container, options)`.
- Multiple simultaneous documents by creating one component instance per container.
- Independent-origin `office-host.html` with per-editor wildcard host support for document-by-document iframe teardown and memory isolation.
- Local conversion through `public/wasm/x2t`.
- OnlyOffice browser runtime under `public/web-apps` and `public/sdkjs`.
- A minimal demo/test host at `/`.
- Integration docs in [docs/integration.md](docs/integration.md) and [docs/integration.zh.md](docs/integration.zh.md).

## Install

```bash
npm install @agentbridges-ai/onlyoffice-browser
```

```ts
import { createOfficeEditor } from '@agentbridges-ai/onlyoffice-browser';
```

The npm package intentionally contains the JS/TS component API only. Deploy `office-host.html` and the OnlyOffice runtime assets from this repository's `public/` directory, a release artifact, or your own CDN on an independent origin.

The runtime assets must be reachable from the editor host origin:

- `/web-apps/` and `/sdkjs/`
- `/wasm/x2t/`
- `/dictionaries/` and `/libs/`
- generated font assets: `/onlyoffice-browser-font-assets.json`, `/sdkjs/common/AllFonts.js`, `/sdkjs/common/Images/fonts_thumbnail*.png`, `/fonts/`, and `/server/FileConverter/bin/font_selection.bin`
- `/document_editor_service_worker.js`, `/plugins.json`, `/themes.json`, and `/reset.html`

The production build uses a compact runtime profile. It keeps the Word, Spreadsheet, and Presentation editors, shared runtime files, `x2t`, `libs`, and `en_US` dictionaries, and removes low-frequency bundled assets such as PDF/Visio SDKs, non-selected dictionaries, bundled help images, and package fonts. It also emits canonical-path packs under `.onlyoffice-runtime-asset-packs/{core,word,cell,slide}`. Deploy `core` plus the document-type packs you support; each pack can be copied to the same editor-host root without rewriting OnlyOffice paths.

## Usage

```ts
import { createOfficeEditor } from '@agentbridges-ai/onlyoffice-browser';

const editor = await createOfficeEditor(document.querySelector('#editor') as HTMLElement, {
  hostUrl: ({ sessionId }) => `https://${sessionId}.office-host.example.com/office-host.html`,
  file,
  fileName: file.name,
  mode: 'edit',
  lang: 'en',
  saveBehavior: 'callback',
  onSave(savedFile) {
    console.log('write this file to your storage target', savedFile.name, savedFile.size);
  },
  onDirtyChange(dirty) {
    console.log('dirty', dirty);
  },
});

// Programmatic saves remain available for automation. User-facing UI should use
// the native OnlyOffice Save button in the editor toolbar.
const saved = await editor.save('DOCX');
editor.setReadonly(true);
editor.setReadonly(false);
editor.destroy();
```

Supported `mode` values:

- `edit`: full editor.
- `readonly`: opens the editor runtime with edit rights disabled; `setReadonly(false)` can restore editing.
- `preview`: upstream OnlyOffice embedded viewer with no editing ribbon; recreate the instance to switch back to editing.

Spellcheck is disabled by default. Pass `spellcheck: true` only when the host app should open documents with spell checking enabled.
Word and presentation documents opened in the editor runtime default to fit-to-width zoom so the page/slide uses the available preview area.
Autosave and force-save are disabled. User-facing saves should go through the native OnlyOffice Save button; `editor.save()` is kept for programmatic integrations and tests.

For multiple documents, create one container and one component instance per document. Prefer wildcard DNS/TLS so each instance gets its own host origin, such as `https://<session>.office-host.example.com/office-host.html`; this lets the corresponding subframe task exit when an individual document closes. In local development, `.localhost` hosts are automatically derived into `host-<session>.localhost`.

## Development

```bash
pnpm install
pnpm run dev
pnpm run lint
pnpm run test
pnpm run test:e2e:save
pnpm run test:e2e:print
pnpm run build:lib
pnpm run pack:dry
pnpm run build
```

`pnpm run test:e2e:save` verifies manual native-save routing for existing local files, new documents, and legacy `.doc`/`.xls`/`.ppt` to OOXML output. `pnpm run test:e2e:print` verifies the native print flow for modern and legacy Word/Excel/PowerPoint files. `pnpm run test:real-load` runs the real-file same-page stress test. See the integration guide for Chrome CDP flags and file arguments.

## Build

```bash
pnpm run build
```

Deploy the generated `dist/` directory and runtime assets as static files on the editor host origin. The parent app must pass that host's `office-host.html` URL as `hostUrl`.

To build optimized runtime assets separately from the demo build:

```bash
npm run assets:build
```

The generated `.onlyoffice-runtime-assets` directory is directly deployable. The `.onlyoffice-runtime-asset-packs` directory contains split `core`, `word`, `cell`, and `slide` packs for CDN or release-artifact workflows. Pass `--types word,slide` or `--dictionaries en_US,en_GB` to `scripts/build-onlyoffice-runtime-assets.mjs` when you need a narrower or broader deployment profile. Preview/edit mode assets are intentionally not split yet; both `embed` and `main` shells stay inside each document-type pack.

## GitHub Actions

- `CI`: lint, TypeScript, unit tests, npm package build/dry-pack, demo production build, and Playwright E2E.
- `Deploy demo`: manual GitHub Pages deployment for the demo. It is manual because the OnlyOffice runtime assets are large.
- `Publish npm`: tag or manual npm release workflow prepared for npm Trusted Publishing.

To enable `Publish npm`, configure npm Trusted Publishing for:

- Package: `@agentbridges-ai/onlyoffice-browser`
- Repository: `agentbridges-ai/onlyoffice-browser`
- Workflow: `publish.yml`
- Environment: leave blank unless the workflow is later changed to use one

The workflow uses GitHub OIDC through `id-token: write`, so it does not need a long-lived npm token once the trusted publisher is configured.

## Browser-Only Boundary

- No OnlyOffice DocumentServer.
- No document upload during normal open/edit/save.
- No `/doc/.../c`, websocket, long-poll, CommandService, or ConvertService dependency.
- Blob URLs live only for the current tab lifetime.
- Collaboration, server-side history, and direct filesystem writes inside the package are out of scope; host apps own `onSave` persistence.

## Fonts

This project does not ship runtime font files. Generate and mount fonts before opening documents: `npm run fonts:generate -- --input /path/to/fonts --output .onlyoffice-font-assets` or `npx onlyoffice-browser-generate-font-assets --input /path/to/fonts --output .onlyoffice-font-assets`, verify with `npm run fonts:verify -- --input .onlyoffice-font-assets` or `npx onlyoffice-browser-verify-font-assets --input .onlyoffice-font-assets`, then start local dev with `npm run dev:fonts`. The generator defaults to a compact `zh-core` set for Simplified Chinese and classic English Office documents; pass `--font-set full` only when broad language coverage is required. See [docs/fonts.md](docs/fonts.md).

## References

- [cryptpad/onlyoffice-editor](https://github.com/cryptpad/onlyoffice-editor)
- [cryptpad/onlyoffice-x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)
- [ONLYOFFICE/Docker-DocumentServer](https://github.com/ONLYOFFICE/Docker-DocumentServer)
- [ONLYOFFICE web-apps](https://github.com/ONLYOFFICE/web-apps)
- [ONLYOFFICE sdkjs](https://github.com/ONLYOFFICE/sdkjs)
- [npm Trusted Publishing](https://docs.npmjs.com/trusted-publishers/)

## License

[AGPL-3.0](LICENSE)
