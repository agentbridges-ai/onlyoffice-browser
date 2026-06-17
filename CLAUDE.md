# CLAUDE.md — Browser Office Editor

## Project Shape

This branch is a browser-only Office preview/edit component, not a full product page. The public API is:

```ts
import { createOfficeEditor } from './src/lib/office-editor';
```

The host application passes a DOM container and owns any multi-document layout. Each editor instance owns its iframe, `Editor.bin` Blob URL, media URLs, mock server, readonly state, and save request.

## Useful Commands

```bash
npm run dev
npm run lint
npm run test
npm run test:e2e
npm run build
```

The repository still declares pnpm in `packageManager`, but this local workspace currently validates with the checked-in `node_modules` and npm because pnpm/corepack are not available here.

## Runtime Boundaries

- No OnlyOffice DocumentServer integration.
- No backend upload or session service.
- No `/doc/.../c`, websocket, long-poll, CommandService, or ConvertService dependency.
- User document bytes, `Editor.bin`, media, and exports stay in the current browser tab as Blob/File data.
- Blob URLs are tab-lifetime only; refresh does not recover private document content.
- Collaboration, server history, and in-place File System Access writes are out of scope for this phase.

## Key Files

- `src/lib/office-editor.ts`: component API and per-instance lifecycle.
- `src/lib/onlyoffice-mock-server.ts`: browser-only adapter for the CryptPad wrapper.
- `src/lib/converter.ts` and `src/lib/document-converter.ts`: singleton x2t WASM conversion layer.
- `src/index.ts`: minimal demo host for manual validation and stress testing.
- `docs/integration.md` and `docs/integration.zh.md`: integration guide.
- `scripts/chrome-cdp-memory-stress.mjs`: Chrome CDP real-load memory stress tool.

Keep OnlyOffice runtime assets under `public/web-apps` and `public/sdkjs` as upstream assets. Cleanup and lifecycle fixes belong in the component shell.
