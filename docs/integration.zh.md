# Browser Office Editor 集成指南

Browser Office Editor 是一个纯浏览器端 Office 预览/编辑组件。宿主应用提供容器 DOM，组件在容器内创建一个轻量 sandbox iframe，指向独立 origin 的 editor host；文档转换、编辑、导出都在这个隔离 host iframe 内完成。

## 部署静态资源

安装组件包：

```bash
pnpm add @agentbridges-ai/onlyoffice-browser
```

npm 主包只包含 JS/TS 组件 API。OnlyOffice runtime 资产需要从本仓库、Release 附件或你自己的 CDN 单独部署。

将 `office-host.html` 和这些运行时资源部署到独立 origin，且该 origin 需要不同于父应用 origin：

- `web-apps/`、`sdkjs/`：OnlyOffice 9.3.0 前端运行时。
- `wasm/x2t/`：浏览器内文档转换引擎。
- `dictionaries/`、`libs/`：字典和 CSV 依赖。
- 生成后的字体资产：`onlyoffice-browser-font-assets.json`、`sdkjs/common/AllFonts.js`、`sdkjs/common/Images/fonts_thumbnail*.png`、`fonts/`、`server/FileConverter/bin/font_selection.bin`。
- `document_editor_service_worker.js`、`plugins.json`、`themes.json`、`reset.html`：OnlyOffice wrapper 需要的根路径文件。

仓库构建默认会生成精简 runtime：保留共享资产以及 Word、Spreadsheet、Presentation 三类编辑器，移除 PDF/Visio SDK、未选中字典、包内字体和大体积 help 资源。也可以直接使用同一套优化 CLI：

```bash
npm run assets:build
```

它会生成可直接部署的 `.onlyoffice-runtime-assets`，以及按 canonical 路径拆分的 `.onlyoffice-runtime-asset-packs/{core,word,cell,slide}`。生产 CDN 工作流可以把 `core` 加实际支持的文档类型 pack 发布到同一个 editor-host 根路径。例如只支持 DOCX/PPTX 的产品只需要发布 `core`、`word`、`slide`，可以省掉 `cell`。每个 pack 内部路径仍是 `/web-apps/...`、`/sdkjs/...`、`/wasm/...` 等，不需要也不应该改写 OnlyOffice 内部资源 URL。

本 profile 暂不拆 preview/edit 模式。每个文档类型 pack 内仍保留上游 `embed` 和 `main` 外壳，因为它们共享渲染/转换资产；现在强拆会带来更高路径和生命周期风险，体积收益相对有限。

创建编辑器时通过 `hostUrl` 传入这个 host 页面。组件会拒绝与父页面同 origin 的 `hostUrl`，因为内存隔离清理依赖独立 iframe origin。

如果同一页面会同时打开多个大文档，推荐为每个 editor 分配一个独立 host origin，而不是所有实例共用同一个 `office-host.example.com`。生产环境可配置 wildcard DNS/TLS，例如 `*.office-host.example.com`，然后用 `hostUrl` resolver 返回每个实例自己的子域名；这样逐个关闭文档时，对应子框架 renderer 可以单独退出。开发环境下，组件会自动把 `.localhost` host 派生为 `host-<session>.localhost`，用于复现同样的逐实例隔离效果。

不需要 OnlyOffice DocumentServer，不需要后端会话服务，也不需要上传用户文档。

## 基础用法

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

`mode` 可选值：

- `edit`：默认编辑运行态。
- `readonly`：使用编辑运行态打开，但立即禁用编辑权；可通过 `setReadonly(false)` 切回编辑。
- `preview`：使用 OnlyOffice 上游 `embedded` viewer 外壳和 `view` 模式，不显示编辑 ribbon；不能原地切回编辑，需要销毁后用 `mode: 'edit'` 重建。

`readonly: true` 等价于 `mode: 'readonly'`，保留用于兼容旧调用。

拼写检查默认关闭。只有宿主应用确实需要默认打开拼写检查时，才传入 `spellcheck: true`。
Word 和演示文稿在编辑器运行态中默认使用适合宽度缩放，让页面/幻灯片优先占满可用预览区域。

## 多文档同时打开

组件不内置标签栏。宿主应用为每个文档创建一个容器，并保存返回的实例：

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

OnlyOffice API script 和 x2t WASM 初始化由各自 host iframe 内部完成；多个实例会各自拥有 iframe 运行上下文，这是同时编辑多个文档的必要成本。若希望“关闭其中一个文档”时 Chrome 任务管理器里的子框架内存同步下降，请使用上面的 per-editor wildcard host origin。

## 打开文档

```ts
// 新建
await createOfficeEditor(container, { hostUrl: officeHostUrl, emptyType: 'docx', fileName: 'New_Document.docx' });

// File / Blob
await createOfficeEditor(container, { hostUrl: officeHostUrl, file, fileName: file.name });

// ArrayBuffer / Uint8Array
await createOfficeEditor(container, { hostUrl: officeHostUrl, buffer, fileName: 'report.xlsx' });

// URL，需 CORS 允许浏览器读取
await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  url: 'https://example.com/report.pptx',
  fileName: 'report.pptx',
  fetchOptions: { headers: { Authorization: `Bearer ${token}` } },
});
```

支持 DOCX、XLSX、PPTX、CSV，也兼容常见旧格式 DOC/XLS/PPT 的打开转换。

## 保存和上传

```ts
const file = await editor.save('XLSX');

await fetch('/api/files/123', {
  method: 'PUT',
  body: file,
});
```

`save(targetExt?)` 返回浏览器内生成的 `File`。常用 `targetExt` 为 `DOCX`、`XLSX`、`PPTX`、`CSV`。只读实例会拒绝保存。本包不会自行持久化文件：宿主应用必须把这个 `File` 写入自己的最终存储位置，例如后端上传接口，或 File System Access 的 `createWritable()` 文件句柄。

OnlyOffice iframe 内部工具栏保存也遵循同一个规则。需要自动持久化编辑结果时，传入 `onSave(file)`：

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

在上游 ONLYOFFICE DocumentServer 部署里，这个持久化角色通常由 `callbackUrl` 后面的 storage service 实现：编辑器上报保存状态，storage service 下载编辑后的文件 URL，写入最终路径，并返回 `{ "error": 0 }`。本包是纯浏览器集成，没有 server callback endpoint，因此集成方必须提供最后的写回步骤。开发时以 [ONLYOFFICE callback handler](https://api.onlyoffice.com/docs/docs-api/usage-api/callback-handler/)、[ONLYOFFICE saving file](https://api.onlyoffice.com/docs/docs-api/get-started/how-it-works/saving-file/)、[ONLYOFFICE/Docker-DocumentServer](https://github.com/ONLYOFFICE/Docker-DocumentServer)、[cryptpad/onlyoffice-editor](https://github.com/cryptpad/onlyoffice-editor) 和 [cryptpad/onlyoffice-x2t-wasm](https://github.com/cryptpad/onlyoffice-x2t-wasm) 为集成参考。

## 只读和关闭

```ts
editor.setReadonly(true);
editor.setReadonly(false);

const state = editor.getState();
console.log(state.mode, state.readonly, state.destroyed);

editor.destroy();
```

只读模式是可恢复的临时编辑权控制，适合在同一个编辑实例中切换查看/编辑。纯预览模式使用上游 OnlyOffice embedded viewer：

```ts
const preview = await createOfficeEditor(container, {
  hostUrl: officeHostUrl,
  file,
  fileName: file.name,
  mode: 'preview',
});
```

纯预览实例不能通过 `setReadonly(false)` 变成编辑实例；需要先 `destroy()`，再用 `mode: 'edit'` 或默认配置重新创建。

`destroy()` 是幂等的，会向隔离 host 发送 `DESTROY`。host 销毁 OnlyOffice wrapper、revoke 本地 Blob URL、清理 worker/socket/timer/media、注销 host origin 下的 service worker/cache，然后把自身导航到轻量 `reset.html`；父页面收到 `HOST_RESET_DONE` 后再把外层 iframe 导航到 `about:blank`、移除 iframe、关闭 `MessagePort` 并清空父侧引用。默认不会刷新父页面；`hardResetOnLastDestroy: true` 仅建议作为最后一个实例关闭后的诊断 fallback。

## 字体资产

本包不包含 runtime 字体文件。打开文档前必须先生成官方 OnlyOffice 字体资产：

```bash
npm run fonts:generate -- --input /absolute/path/to/fonts --output .onlyoffice-font-assets
```

如果是在应用项目中使用已发布的 npm 包，使用包内 CLI：

```bash
npx onlyoffice-browser-generate-font-assets --input /absolute/path/to/fonts --output .onlyoffice-font-assets
```

然后由 editor host 提供这些资产。本地开发示例：

```bash
npm run fonts:verify -- --input .onlyoffice-font-assets
npm run dev:fonts
```

这会把生成出的 `AllFonts.js`、字体缩略图、字体文件和 `font_selection.bin` 作为静态资产提供给编辑器。缺少字体资产会被视为配置错误；本包不提供默认 fallback。精简字体集会保留完整字体注册表用于文档解析，并按生成资产里的可见字体清单过滤 UI 字体下拉。详见 [fonts.zh.md](fonts.zh.md)。

## 隐私边界

- 文档内容通过 `postMessage`/transferables 从父页面传入 host，不会上传到任何服务。
- 正常打开、编辑、保存只应请求 host origin 静态资源和 host 内创建的本地 `blob:` URL。
- 不使用 `/doc/.../c`、websocket、long-poll、CommandService、ConvertService。
- Blob URL 只在当前标签页生命周期内有效；刷新页面不会恢复私有文件。
- 本组件不提供协同编辑、服务端历史或原位写回。

## Network 验收清单

集成后用浏览器 Network 面板检查：

- `web-apps/apps/api/documents/api.js` 从配置的 host origin 加载。
- `wasm/x2t/x2t.js` 和 `x2t.wasm` 在 host iframe 内初始化。
- 多实例关闭后父页面中没有 editor host iframe，也没有嵌套的 `iframe[name="frameEditor"]`。
- 多实例并发时，推荐看到每个文档独立的 host 子框架任务，例如 `子框架: https://<session>.office-host.example.com/`。
- 没有 DocumentServer 类请求：`/doc/.../c`、`CommandService`、`ConvertService`、`coauthor`、`sockjs`、`websocket`。
- 关闭后 Chrome heap、DOM counters、RSS 不持续随次数增长。

## 真实文件负载测试

启动 Demo 和带 CDP 的 Chrome 后，运行同页多文档真实负载测试：

```bash
npm run test:real-load -- \
  --browser-ws=ws://localhost:9222/devtools/browser \
  --url=http://127.0.0.1:5173/ \
  --host-url=http://host.localhost:5173/office-host.html \
  --file=/path/to/real.docx \
  --file=/path/to/real.pptx \
  --file=/path/to/real.xls
```

默认每轮同时打开 3 个真实文件，并混合覆盖 `edit`、`readonly`、`preview` 三种模式；脚本会等待每个实例触发 `onDocumentReady` 后再停留 2 秒，然后通过 Demo 的 Close All 关闭，循环 60 次并输出 JSON/CSV 到 `test-results/memory/`。如果需要改变并发数，可传 `--batch-size=N`，真实文件模式下需要至少提供 N 个 `--file`。

也可以通过 `--modes=preview` 单独压测上游 embedded 纯预览模式，或通过 `--modes=edit,readonly,preview --batch-size=3` 明确做三模式混合批次。
