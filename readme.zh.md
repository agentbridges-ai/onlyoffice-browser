# Browser Office Editor

[![CI](https://github.com/agentbridges-ai/onlyoffice-browser/actions/workflows/ci.yml/badge.svg)](https://github.com/agentbridges-ai/onlyoffice-browser/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@agentbridges-ai/onlyoffice-browser.svg)](https://www.npmjs.com/package/@agentbridges-ai/onlyoffice-browser)
[![license](https://img.shields.io/github/license/agentbridges-ai/onlyoffice-browser.svg)](LICENSE)

基于 OnlyOffice 9.3 和 `onlyoffice-x2t-wasm` 的纯浏览器端 Office 预览/编辑组件。

组件在宿主传入的 DOM 容器内打开 DOCX、XLSX、PPTX、CSV。转换、编辑、保存、导出都在独立 origin 的 sandbox editor host iframe 中完成，不需要 OnlyOffice DocumentServer、不需要后端会话、不上传文档、不需要用户账号。

## 仓库提供什么

- JS/TS API：`createOfficeEditor(container, options)`。
- 一页多文档：宿主为每个文档创建一个容器和一个组件实例。
- 独立 origin 的 `office-host.html`，支持 per-editor wildcard host，用于逐个文档销毁 iframe 和隔离内存。
- `public/wasm/x2t` 本地转换链路。
- `public/web-apps`、`public/sdkjs` OnlyOffice 浏览器运行时。
- `/` 极简 demo/test 宿主页面。
- 集成文档：[docs/integration.zh.md](docs/integration.zh.md) 和 [docs/integration.md](docs/integration.md)。

## 安装

```bash
npm install @agentbridges-ai/onlyoffice-browser
```

```ts
import { createOfficeEditor } from '@agentbridges-ai/onlyoffice-browser';
```

npm 主包只包含 JS/TS 组件 API。`office-host.html` 和 OnlyOffice runtime 资产需要从本仓库 `public/`、Release 附件或你自己的 CDN 单独部署到独立 origin。

运行时资源需要能从 editor host origin 访问：

- `/web-apps/` 和 `/sdkjs/`
- `/wasm/x2t/`
- `/dictionaries/`、`/libs/`
- 生成后的字体资产：`/onlyoffice-browser-font-assets.json`、`/sdkjs/common/AllFonts.js`、`/sdkjs/common/Images/fonts_thumbnail*.png`、`/fonts/`、`/server/FileConverter/bin/font_selection.bin`
- `/document_editor_service_worker.js`、`/plugins.json`、`/themes.json`、`/reset.html`

## 使用

```ts
import { createOfficeEditor } from '@agentbridges-ai/onlyoffice-browser';

const editor = await createOfficeEditor(document.querySelector('#editor') as HTMLElement, {
  hostUrl: ({ sessionId }) => `https://${sessionId}.office-host.example.com/office-host.html`,
  file,
  fileName: file.name,
  mode: 'edit',
  lang: 'zh',
  onSave(savedFile) {
    console.log(savedFile.name, savedFile.size);
  },
});

const saved = await editor.save('DOCX');
editor.setReadonly(true);
editor.setReadonly(false);
editor.destroy();
```

`mode` 可选值：

- `edit`：完整编辑器。
- `readonly`：编辑运行时，但禁用编辑权；可通过 `setReadonly(false)` 恢复编辑。
- `preview`：OnlyOffice 上游 embedded viewer，不显示编辑 ribbon；需要销毁并重建实例才能切回编辑。

多文档同时打开时，宿主为每个文档创建一个容器和一个组件实例即可。推荐使用 wildcard DNS/TLS 给每个实例分配独立 host origin，例如 `https://<session>.office-host.example.com/office-host.html`；这样逐个关闭文档时，对应子框架任务可以独立退出。开发环境下，`.localhost` host 会自动派生为 `host-<session>.localhost`。

## 开发

```bash
pnpm install
pnpm run dev
pnpm run lint
pnpm run test
pnpm run build:lib
pnpm run pack:dry
pnpm run build
```

`pnpm run test:real-load` 用于真实文件同页压力测试。Chrome CDP 参数和文件参数见集成文档。

## 构建

```bash
pnpm run build
```

将生成的 `dist/` 和运行时资源部署到 editor host origin。父应用创建实例时必须通过 `hostUrl` 传入这个 host 的 `office-host.html` 地址。

## GitHub Actions

- `CI`：lint、TypeScript、单元测试、npm package build/dry-pack、demo 生产构建、Playwright E2E。
- `Deploy demo`：手动部署 GitHub Pages demo。OnlyOffice runtime 资产较大，所以不在每次 main push 自动上传。
- `Publish npm`：基于 tag 或手动触发的 npm 发布 workflow，已按 npm Trusted Publishing 准备。

启用 `Publish npm` 时，在 npm package 设置里配置 Trusted Publisher：

- Package：`@agentbridges-ai/onlyoffice-browser`
- Repository：`agentbridges-ai/onlyoffice-browser`
- Workflow：`publish.yml`
- Environment：留空，除非之后 workflow 显式增加 environment

配置完成后 workflow 会通过 GitHub OIDC 发布，不再需要长期 npm token。

## 浏览器内边界

- 不使用 OnlyOffice DocumentServer。
- 正常打开、编辑、保存不会上传用户文档。
- 不依赖 `/doc/.../c`、websocket、long-poll、CommandService、ConvertService。
- Blob URL 只在当前标签页生命周期内有效。
- 协同编辑、服务端历史、原位写回不在本轮范围内。

## 字体

本项目不内置 runtime 字体文件。打开文档前必须先生成并挂载字体资产：`npm run fonts:generate -- --input /path/to/fonts --output .onlyoffice-font-assets` 或 `npx onlyoffice-browser-generate-font-assets --input /path/to/fonts --output .onlyoffice-font-assets`，再用 `npm run fonts:verify -- --input .onlyoffice-font-assets` 或 `npx onlyoffice-browser-verify-font-assets --input .onlyoffice-font-assets` 验证，本地开发用 `npm run dev:fonts` 启动。生成器默认使用精简的 `zh-core` 字体集，面向简体中文和经典英文 Office 文档；只有确实需要广泛语言覆盖时再传 `--font-set full`。详见 [docs/fonts.zh.md](docs/fonts.zh.md)。

## 参考

- [cryptpad/onlyoffice-editor](https://github.com/cryptpad/onlyoffice-editor)
- [cryptpad/onlyoffice-x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm)
- [ONLYOFFICE web-apps](https://github.com/ONLYOFFICE/web-apps)
- [ONLYOFFICE sdkjs](https://github.com/ONLYOFFICE/sdkjs)
- [npm Trusted Publishing](https://docs.npmjs.com/trusted-publishers/)

## 许可证

[AGPL-3.0](LICENSE)
