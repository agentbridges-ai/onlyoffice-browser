# Browser Office Editor

A browser-only Office preview/edit component powered by OnlyOffice 9.3 and `onlyoffice-x2t-wasm`.

The component opens DOCX, XLSX, PPTX, and CSV documents inside a host-provided DOM container. Conversion, editing, saving, and export all happen in the current browser tab. It does not require OnlyOffice DocumentServer, backend sessions, document uploads, or user accounts.

## What This Repo Provides

- JS/TS API: `createOfficeEditor(container, options)`.
- Multiple simultaneous documents by creating one component instance per container.
- Local conversion through `public/wasm/x2t`.
- OnlyOffice browser runtime under `public/web-apps` and `public/sdkjs`.
- A small demo/test host at `/`.
- Integration docs in [docs/integration.md](docs/integration.md).

## Quick Start

```bash
pnpm install
pnpm run dev
```

```ts
import { createOfficeEditor } from './src/lib/office-editor';

const editor = await createOfficeEditor(document.querySelector('#editor') as HTMLElement, {
  file,
  fileName: file.name,
});

const saved = await editor.save('DOCX');
editor.setReadonly(true);  // reversible readonly; use mode: 'preview' for the upstream viewer without the editing ribbon
editor.destroy();
```

## Build

```bash
pnpm run build
```

Deploy the generated `dist/` directory as static files. Keep the runtime assets from `public/` on the same origin as the app.

## Browser-Only Boundary

- No OnlyOffice DocumentServer.
- No document upload during normal open/edit/save.
- No `/doc/.../c`, websocket, long-poll, CommandService, or ConvertService dependency.
- Blob URLs live only for the current tab lifetime.
- Collaboration, server-side history, and in-place filesystem writeback are out of scope.

## Fonts

This project does not include proprietary fonts. See [docs/fonts.md](docs/fonts.md) for font asset guidance.

## References

- [cryptpad/onlyoffice-editor](https://github.com/cryptpad/onlyoffice-editor)
- [cryptpad/onlyoffice-x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)
- [ONLYOFFICE web-apps](https://github.com/ONLYOFFICE/web-apps)
- [ONLYOFFICE sdkjs](https://github.com/ONLYOFFICE/sdkjs)

## License

[AGPL-3.0](LICENSE)
