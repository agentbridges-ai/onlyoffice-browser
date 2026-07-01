# Browser Office Editor Integration Guide

Browser Office Editor is a browser-only Office preview/edit component. The host application provides a DOM container; the component creates a lightweight iframe that points at an independent-origin editor host. Conversion, editing, and export happen inside that isolated host iframe. The independent origin is the isolation boundary. Do not add an outer `sandbox` attribute: the editor host relies on same-origin access to its nested print iframe for the native PDF print flow, and the runtime removes a sandbox attribute if an integration mutates it back onto the host iframe.

## Static Assets

Install the component package:

```bash
pnpm add @agentbridges-ai/onlyoffice-browser
```

The npm package intentionally contains only the JS/TS component API. Deploy the runtime assets separately from this repository, a release artifact, or your own CDN.

Deploy `office-host.html` and these runtime assets from the same independent origin, separate from the parent application origin:

- `web-apps/`, `sdkjs/`: OnlyOffice 9.3.0 browser runtime.
- `wasm/x2t/`: browser-side document conversion engine.
- `dictionaries/`, `libs/`: dictionaries and CSV dependencies.
- generated font assets: `onlyoffice-browser-font-assets.json`, `sdkjs/common/AllFonts.js`, `sdkjs/common/Images/fonts_thumbnail*.png`, `fonts/`, and `server/FileConverter/bin/font_selection.bin`.
- `document_editor_service_worker.js`, `plugins.json`, `themes.json`, `reset.html`: root files expected by the OnlyOffice wrapper.

The repository build produces a compact runtime by default. It keeps shared assets plus the Word, Spreadsheet, and Presentation editor families, and removes bundled PDF/Visio SDKs, non-selected dictionaries, package fonts, and large help resources. The same optimization CLI can be used directly:

```bash
npm run assets:build
```

This writes a directly deployable `.onlyoffice-runtime-assets` directory and split canonical-path packs under `.onlyoffice-runtime-asset-packs/{core,word,cell,slide}`. For production CDN workflows, publish `core` plus the document-type packs your product supports to the same editor-host root. For example, a DOCX/PPTX-only product can publish `core`, `word`, and `slide`, omitting `cell`. The paths inside each pack remain `/web-apps/...`, `/sdkjs/...`, `/wasm/...`, etc.; do not rewrite OnlyOffice internal asset URLs.

Preview/edit mode assets are not split in this profile. Each document-type pack keeps both upstream `embed` and `main` shells, because they share renderer/conversion assets and splitting them creates more path and lifecycle risk than size benefit for now.

Pass that host page as `hostUrl`. The component rejects same-origin `hostUrl` values because the memory-isolation cleanup relies on an independent iframe origin.

If a page opens multiple large documents at the same time, prefer one host origin per editor instead of sharing a single `office-host.example.com` origin. In production, configure wildcard DNS/TLS such as `*.office-host.example.com`, then pass a `hostUrl` resolver that returns a session-specific subdomain. This lets Chrome retire each editor's subframe renderer when that document closes. In local development, `.localhost` host URLs are automatically derived into `host-<session>.localhost` to emulate the same per-instance isolation.

No OnlyOffice DocumentServer, backend session service, or document upload endpoint is required.

## Basic Usage

```ts
import { createOfficeEditor } from '@agentbridges-ai/onlyoffice-browser';

const container = document.querySelector('#editor') as HTMLElement;
const officeHostUrl = ({ sessionId }: { sessionId: string }) =>
  `https://${sessionId}.office-host.example.com/office-host.html`;

const editor = await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  file: fileInput.files![0],
  fileName: fileInput.files![0].name,
  mode: 'edit',
  readonly: false,
  saveBehavior: 'callback',
  onReady(instance) {
    console.log('ready', instance.getState());
  },
  onSave(file) {
    console.log('write this file to your storage target', file.name, file.size);
  },
  onDirtyChange(dirty) {
    console.log('dirty', dirty);
  },
  onError(error) {
    console.error(error);
  },
});
```

`mode` values:

- `edit`: default editable runtime.
- `readonly`: opens the editable runtime and immediately disables edit rights; `setReadonly(false)` can restore editing.
- `preview`: uses the upstream OnlyOffice `embedded` viewer shell with `view` mode, so the editing ribbon is not shown; it cannot be switched to editing in place. Destroy and recreate with `mode: 'edit'` instead.

`readonly: true` is equivalent to `mode: 'readonly'` and remains supported for compatibility.

Spellcheck is disabled by default. Pass `spellcheck: true` only when the host app should open documents with spell checking enabled.
Word and presentation documents opened in the editor runtime default to fit-to-width zoom so the page/slide uses the available preview area.

## Multiple Documents

The component does not include tabs. The host creates one container per document and stores the returned instances:

```ts
const editors = await Promise.all(
  files.map((file) => {
    const container = document.createElement('div');
    container.className = 'editor-pane';
    document.querySelector('#workspace')!.appendChild(container);
    return createOfficeEditor(container, { hostUrl: officeHostUrl, file, fileName: file.name });
  }),
);

await editors[0].save('DOCX');
editors[1].setReadonly(true);
editors[2].destroy();
```

The OnlyOffice API script and x2t WASM initialization happen inside each host iframe. Each editor owns its iframe runtime context, which is required for simultaneous editing. To make Chrome Task Manager memory drop as individual documents close, use a per-editor wildcard host origin as shown above.

## Opening Documents

```ts
// New document. The native OnlyOffice Save button downloads by default.
await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  emptyType: 'docx',
  fileName: 'New_Document.docx',
  saveBehavior: 'download',
});

// File / Blob. The native OnlyOffice Save button calls onSave.
await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  file,
  fileName: file.name,
  saveBehavior: 'callback',
  onSave: async (savedFile) => {
    await uploadOrWriteBack(savedFile);
  },
});

// ArrayBuffer / Uint8Array
await createOfficeEditor(container, { hostUrl: officeHostUrl, buffer, fileName: 'report.xlsx' });

// URL; the source must be CORS-readable by the browser
await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  url: 'https://example.com/report.pptx',
  fileName: 'report.pptx',
  fetchOptions: { headers: { Authorization: `Bearer ${token}` } },
});
```

DOCX, XLSX, PPTX, and CSV are the primary targets. Common legacy DOC/XLS/PPT files can be opened through conversion. Treat those legacy formats as compatibility inputs: editable output is normalized to OOXML (`.docx`, `.xlsx`, `.pptx`) instead of trying to write the old binary Office format back in place.

## Saving

```ts
const file = await editor.save('XLSX');

await fetch('/api/files/123', {
  method: 'PUT',
  body: file,
});
```

`save(targetExt?)` returns a browser-generated `File`. Common targets are `DOCX`, `XLSX`, `PPTX`, and `CSV`. Read-only instances reject save requests. This API remains available for programmatic integrations and tests. User-facing UI should use the native OnlyOffice Save button in the editor toolbar, not a duplicate host-side Save button.

Requests to save legacy `DOC`, `XLS`, or `PPT` are converted to the matching OOXML result (`DOCX`, `XLSX`, `PPTX`). Host apps that open a local legacy file should treat the returned `File.name` as authoritative and update their storage path or metadata accordingly, rather than writing OOXML bytes into a `.doc`, `.xls`, or `.ppt` path.

Autosave and force-save are disabled in the embedded OnlyOffice config, and co-editing is forced to `strict` mode so upstream fast-mode autosave cannot turn itself back on. Only a manual native Save should persist or download edited content.

The package exports edited bytes by calling the browser runtime's native bin export and converting that bin with the bundled x2t WASM converter. It does not call `downloadAs()` for local persistence, so callback saves do not open a browser download dialog.

Document resources must travel with that native bin. Upstream DocumentServer downloads the whole document storage directory into the converter `source` folder before running FileConverter, so `Editor.bin`, changes, and sidecar files such as `media/...` are available to x2t together. This browser package has no server storage directory, so it keeps the media object URL map from open/insert operations and materializes those resources back into the x2t `/working/media` folder before Save, Print, or Download as conversion. Do not change export code to pass only the native bin; that can produce OOXML packages whose relationships still reference images while `word/media/*`, `xl/media/*`, or `ppt/media/*` files are missing.

`saveBehavior` controls what happens after native export:

- `auto` (default): existing `file`/`buffer`/`url` sources use `onSave`; missing `onSave` rejects. `emptyType` new documents download unless `onSave` returns `true`.
- `callback`: require `onSave` and wait for it. Use this for local files that must be written back through File System Access.
- `download`: trigger a browser download. Use this for blank documents that do not yet have a storage path.

`onSave(file)` is called by the native Save path and by `editor.save()`, and the save acknowledgement waits for it. If `onSave` throws or rejects, the proxy keeps the document dirty and the native Save is acknowledged as failed:

```ts
await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  file,
  fileName: file.name,
  onSave: async (savedFile) => {
    const writable = await fileHandle.createWritable();
    await writable.write(savedFile);
    await writable.close();
  },
});
```

The native OnlyOffice "All changes saved" status is not a host persistence signal in this browser-only integration. In upstream ONLYOFFICE DocumentServer deployments that role is normally implemented by the storage service behind `callbackUrl`: the editor reports save status, the storage service downloads the edited file URL, writes it to the final path, and returns `{ "error": 0 }`. This package has no server callback endpoint, so the integrator must provide the final write step. Use [ONLYOFFICE callback handler](https://api.onlyoffice.com/docs/docs-api/usage-api/callback-handler/), [ONLYOFFICE saving file](https://api.onlyoffice.com/docs/docs-api/get-started/how-it-works/saving-file/), [ONLYOFFICE/DocumentServer](https://github.com/ONLYOFFICE/DocumentServer), [ONLYOFFICE/Docker-DocumentServer](https://github.com/ONLYOFFICE/Docker-DocumentServer), [cryptpad/onlyoffice-editor](https://github.com/cryptpad/onlyoffice-editor), and [cryptpad/onlyoffice-x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm) as the integration references.

## Printing

Use the native OnlyOffice Print button. Upstream DocumentServer prints from the editor iframe by calling `asc_Print()`, generating a server PDF URL, loading that URL into the hidden `#id-print-frame`, and calling `iframe.contentWindow.print()`. In the official server implementation, the URL is a same-origin `/printfile/:docid/:filename` response with `Content-Type: application/pdf` and `Content-Disposition: inline`; the filename is intentionally included in both the URL path and disposition because Chrome uses the resource name when saving the print result.

The browser runtime mirrors that flow without a server-side `printfile` endpoint: it provides the parent `APP.printPdf` bridge expected by the editor, asks the editor iframe for its native print renderer stream, converts that raw non-base64 stream through x2t `bin2pdf` (`m_bIsNoBase64=true`), stores the resulting PDF in a Cache API entry on the editor host origin, and returns a temporary PDF URL under `/__onlyoffice-browser-print__/.../<file>.pdf?filename=<file>.pdf` on that same editor host origin. The editor host service worker serves that PDF back to OnlyOffice's built-in hidden `#id-print-frame` with PDF response headers matching the official shape. Keep the URL resource name, `Content-Disposition` filename, and PDF document metadata title aligned; Chrome's print/PDF viewer path can derive the "Save as PDF" default name from the PDF title metadata, so the runtime appends an incremental PDF info dictionary with a UTF-16BE `/Title` before caching the generated PDF.

This matches the browser-side constraints for isolated hosts: `window.print()` prints the current loaded document, iframe `contentWindow` access is governed by same-origin policy, and modern `blob:` URLs are subject to storage partitioning/navigation restrictions. Do not use an outer sandbox, even with `allow-same-origin` and `allow-modals`; Chrome can still classify the nested PDF print frame as inaccessible for this native flow. The packaged OnlyOffice runtime is patched at build time to keep the hidden print iframe, retry `contentWindow.print()` briefly while the PDF document becomes script-accessible, and suppress `window.open()` / `downloadAs()` fallbacks. Do not rasterize the document through images, canvas, or PDF.js for printing.

Reference docs:

- [MDN `Window.print()`](https://developer.mozilla.org/en-US/docs/Web/API/Window/print)
- [MDN `HTMLIFrameElement.contentWindow`](https://developer.mozilla.org/en-US/docs/Web/API/HTMLIFrameElement/contentWindow)
- [MDN same-origin policy](https://developer.mozilla.org/en-US/docs/Web/Security/Defenses/Same-origin_policy)
- [MDN `blob:` URLs and storage partitioning](https://developer.mozilla.org/en-US/docs/Web/URI/Reference/Schemes/blob)
- [MDN service workers](https://developer.mozilla.org/en-US/docs/Web/API/Service_Worker_API/Using_Service_Workers)
- [Chrome sandboxed modal dialogs sample](https://googlechrome.github.io/samples/block-modal-dialogs-sandboxed-iframe/index.html)
- [ONLYOFFICE/web-apps document print controller](https://github.com/ONLYOFFICE/web-apps/blob/master/apps/documenteditor/main/app/controller/Main.js)
- [ONLYOFFICE/web-apps print settings controller](https://github.com/ONLYOFFICE/web-apps/blob/master/apps/documenteditor/main/app/controller/Print.js)

Printing is intentionally independent from saving: it does not call `downloadAs()`, does not write through `onSave`, and does not change dirty state. The print E2E matrix covers `xlsx`, `xls`, `docx`, `doc`, `pptx`, and `ppt`.

## Readonly and Cleanup

```ts
editor.setReadonly(true);
editor.setReadonly(false);

const state = editor.getState();
console.log(state.mode, state.readonly, state.destroyed);

editor.destroy();
```

Readonly mode is a reversible edit-rights toggle for the same editor instance. Pure preview mode uses the upstream OnlyOffice embedded viewer:

```ts
const preview = await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  file,
  fileName: file.name,
  mode: 'preview',
});
```

Preview instances cannot become editable through `setReadonly(false)`. Destroy them and recreate with `mode: 'edit'` or the default options when editing is needed.

`destroy()` is idempotent. It sends `DESTROY` to the isolated host. The host destroys the OnlyOffice wrapper, revokes local Blob URLs, clears worker/socket/timer/media resources, unregisters host-origin service workers/caches, and then navigates itself to the lightweight `reset.html`; the parent removes the outer iframe only after receiving `HOST_RESET_DONE`, then navigates it to `about:blank`, closes the `MessagePort`, and clears parent references. It does not reload the parent page by default. Pass `hardResetOnLastDestroy: true` only as a diagnostic fallback after the last instance closes.

## Font Assets

The package does not include runtime font files. Generate official OnlyOffice font assets before opening documents:

```bash
npm run fonts:generate -- --input /absolute/path/to/fonts --output .onlyoffice-font-assets
```

When using the published package from an application repository, use the package CLI instead:

```bash
npx onlyoffice-browser-generate-font-assets --input /absolute/path/to/fonts --output .onlyoffice-font-assets
```

Then serve them from the editor host. For local dev:

```bash
npm run fonts:verify -- --input .onlyoffice-font-assets
npm run dev:fonts
```

This serves generated `AllFonts.js`, font thumbnails, font files, and `font_selection.bin` as static assets. Missing font assets are treated as configuration errors; there is no package fallback. See [fonts.md](fonts.md).

## Privacy Boundary

- Document content is transferred from the parent to the host with `postMessage`/transferables; it is not uploaded to any service.
- Normal open/edit/save flows request host-origin static assets and local host-created `blob:` URLs only.
- The component does not use `/doc/.../c`, websocket, long-poll, CommandService, or ConvertService.
- Blob URLs are valid only for the current tab lifetime; refresh does not restore private files.
- Collaboration, server-side history, and in-place filesystem writeback are out of scope.

## Network Acceptance Checklist

Use the browser Network panel after integration:

- `web-apps/apps/api/documents/api.js` loads from the configured host origin.
- `wasm/x2t/x2t.js` and `x2t.wasm` initialize inside the host iframe.
- After closing instances, the parent has no editor host iframe and no nested `iframe[name="frameEditor"]`.
- With concurrent instances, prefer seeing one independent host subframe task per document, for example `Subframe: https://<session>.office-host.example.com/`.
- No DocumentServer-like requests appear: `/doc/.../c`, `CommandService`, `ConvertService`, `coauthor`, `sockjs`, `websocket`.
- Closed-state Chrome heap, DOM counters, and RSS do not grow continuously with repeated open/close cycles.

## Real-File Load Test

After starting the demo and a Chrome instance with CDP enabled, run the same-page multi-document real-load test:

```bash
npm run test:real-load -- \
  --browser-ws=ws://localhost:9222/devtools/browser \
  --url=http://127.0.0.1:5173/ \
  --host-url=http://host.localhost:5173/office-host.html \
  --file=/path/to/real.docx \
  --file=/path/to/real.pptx \
  --file=/path/to/real.xls
```

By default each cycle opens 3 real files simultaneously and covers `edit`, `readonly`, and `preview` in the same batch; it waits for every instance to fire `onDocumentReady`, then waits 2 seconds, closes them through the demo Close All control, repeats 60 times, and writes JSON/CSV reports to `test-results/memory/`. Pass `--batch-size=N` to change concurrency; real-file mode requires at least N `--file` arguments.

Use `--modes=preview` to stress only the upstream embedded preview shell, or `--modes=edit,readonly,preview --batch-size=3` for explicit mixed-mode batches.
