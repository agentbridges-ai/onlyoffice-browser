# Browser Office Editor

基于 OnlyOffice 9.3 和 `onlyoffice-x2t-wasm` 的纯浏览器端 Office 预览/编辑组件。

组件在宿主传入的 DOM 容器内打开 DOCX、XLSX、PPTX、CSV。转换、编辑、保存、导出都在当前浏览器标签页中完成，不需要 OnlyOffice DocumentServer、不需要后端会话、不上传文档、不需要用户账号。

## 仓库提供什么

- JS/TS API：`createOfficeEditor(container, options)`。
- 一页多文档：宿主为每个文档创建一个容器和一个组件实例。
- `public/wasm/x2t` 本地转换链路。
- `public/web-apps`、`public/sdkjs` OnlyOffice 浏览器运行时。
- `/` 极简 demo/test 宿主页面。
- 集成文档：[docs/integration.zh.md](docs/integration.zh.md)。

## 快速开始

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
editor.setReadonly(true);  // 可恢复的临时只读；无编辑 ribbon 的上游预览使用 mode: 'preview'
editor.destroy();
```

## 构建

```bash
pnpm run build
```

将生成的 `dist/` 作为静态文件部署即可。`public/` 中的运行时资源需要与应用保持同源。

## 浏览器内边界

- 不使用 OnlyOffice DocumentServer。
- 正常打开、编辑、保存不会上传用户文档。
- 不依赖 `/doc/.../c`、websocket、long-poll、CommandService、ConvertService。
- Blob URL 只在当前标签页生命周期内有效。
- 协同编辑、服务端历史、原位写回不在本轮范围内。

## 字体

本项目不包含专有字体。字体资源说明见 [docs/fonts.zh.md](docs/fonts.zh.md)。

## 参考

- [cryptpad/onlyoffice-editor](https://github.com/cryptpad/onlyoffice-editor)
- [cryptpad/onlyoffice-x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)
- [ONLYOFFICE web-apps](https://github.com/ONLYOFFICE/web-apps)
- [ONLYOFFICE sdkjs](https://github.com/ONLYOFFICE/sdkjs)

## 许可证

[AGPL-3.0](LICENSE)
