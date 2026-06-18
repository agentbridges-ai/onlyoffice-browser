# Browser Office Editor Integration Guide

Browser Office Editor is a browser-only Office preview/edit component. The host application provides a DOM container; the component creates a lightweight sandbox iframe that points at an independent-origin editor host. Conversion, editing, and export happen inside that isolated host iframe.

## Static Assets

Install the component package:

```bash
pnpm add @agentbridges-ai/onlyoffice-browser
```

The npm package intentionally contains only the JS/TS component API. Deploy the runtime assets separately from this repository, a release artifact, or your own CDN.

Deploy `office-host.html` and these runtime assets from the same independent origin, separate from the parent application origin:

- `web-apps/`, `sdkjs/`: OnlyOffice 9.3.0 browser runtime.
- `wasm/x2t/`: browser-side document conversion engine.
- `fonts/`, `dictionaries/`, `libs/`: fonts, dictionaries, and CSV dependencies.
- `document_editor_service_worker.js`, `plugins.json`, `themes.json`, `reset.html`: root files expected by the OnlyOffice wrapper.

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
  onReady(instance) {
    console.log('ready', instance.getState());
  },
  onSave(file) {
    console.log('saved', file.name, file.size);
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
// New document
await createOfficeEditor(container, { hostUrl: officeHostUrl, emptyType: 'docx', fileName: 'New_Document.docx' });

// File / Blob
await createOfficeEditor(container, { hostUrl: officeHostUrl, file, fileName: file.name });

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

DOCX, XLSX, PPTX, and CSV are the primary targets. Common legacy DOC/XLS/PPT files can be opened through conversion.

## Saving

```ts
const file = await editor.save('XLSX');

await fetch('/api/files/123', {
  method: 'PUT',
  body: file,
});
```

`save(targetExt?)` returns a browser-generated `File`. Common targets are `DOCX`, `XLSX`, `PPTX`, and `CSV`. Read-only instances reject save requests.

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
