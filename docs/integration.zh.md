# Browser Office Editor 集成指南

Browser Office Editor 是一个纯浏览器端 Office 预览/编辑组件。宿主应用提供容器 DOM，组件在容器内创建 OnlyOffice iframe；文档转换、编辑、导出都在当前浏览器标签页内完成。

## 部署静态资源

安装组件包：

```bash
pnpm add @agentbridges-ai/onlyoffice-browser
```

npm 主包只包含 JS/TS 组件 API。OnlyOffice runtime 资产需要从本仓库、Release 附件或你自己的 CDN 单独部署。

将构建产物和 `public` 中的运行时资源部署到同一 origin：

- `web-apps/`、`sdkjs/`：OnlyOffice 9.3.0 前端运行时。
- `wasm/x2t/`：浏览器内文档转换引擎。
- `fonts/`、`dictionaries/`、`libs/`：字体、字典和 CSV 依赖。
- `document_editor_service_worker.js`、`plugins.json`、`themes.json`、`reset.html`：OnlyOffice wrapper 需要的根路径文件。

不需要 OnlyOffice DocumentServer，不需要后端会话服务，也不需要上传用户文档。

## 基础用法

```ts
import { createOfficeEditor } from '@agentbridges-ai/onlyoffice-browser';

const container = document.querySelector('#editor') as HTMLElement;

const editor = await createOfficeEditor(container, {
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

## 多文档同时打开

组件不内置标签栏。宿主应用为每个文档创建一个容器，并保存返回的实例：

```ts
const editors = await Promise.all(
  files.map((file) => {
    const container = document.createElement('div');
    container.className = 'editor-pane';
    document.querySelector('#workspace')!.appendChild(container);
    return createOfficeEditor(container, { file, fileName: file.name });
  }),
);

await editors[0].save('DOCX');
editors[1].setReadonly(true);
editors[2].destroy();
```

OnlyOffice API script 和 x2t WASM 初始化由组件内部做 singleton 复用；多个实例会各自拥有 iframe 运行上下文，这是同时编辑多个文档的必要成本。

## 打开文档

```ts
// 新建
await createOfficeEditor(container, { emptyType: 'docx', fileName: 'New_Document.docx' });

// File / Blob
await createOfficeEditor(container, { file, fileName: file.name });

// ArrayBuffer / Uint8Array
await createOfficeEditor(container, { buffer, fileName: 'report.xlsx' });

// URL，需 CORS 允许浏览器读取
await createOfficeEditor(container, {
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

`save(targetExt?)` 返回浏览器内生成的 `File`。常用 `targetExt` 为 `DOCX`、`XLSX`、`PPTX`、`CSV`。只读实例会拒绝保存。

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
  file,
  fileName: file.name,
  mode: 'preview',
});
```

纯预览实例不能通过 `setReadonly(false)` 变成编辑实例；需要先 `destroy()`，再用 `mode: 'edit'` 或默认配置重新创建。

`destroy()` 是幂等的，会取消未完成保存、销毁 OnlyOffice wrapper、移除 iframe、revoke `Editor.bin`/media/source Blob URL、注销本实例媒体映射并清空容器 DOM。默认不会刷新页面；如需在最后一个实例关闭后进行同标签硬重置，创建实例时传 `hardResetOnLastDestroy: true`。

## 隐私边界

- 文档内容、`Editor.bin`、媒体文件和导出文件不会上传到任何服务。
- 正常打开、编辑、保存只应请求同源静态资源和本地 `blob:` URL。
- 不使用 `/doc/.../c`、websocket、long-poll、CommandService、ConvertService。
- Blob URL 只在当前标签页生命周期内有效；刷新页面不会恢复私有文件。
- 本组件不提供协同编辑、服务端历史或原位写回。

## Network 验收清单

集成后用浏览器 Network 面板检查：

- `web-apps/apps/api/documents/api.js` 只插入一次。
- `wasm/x2t/x2t.js` 和 `x2t.wasm` 只初始化一次，后续实例复用。
- 多实例关闭后页面中 `iframe[name="frameEditor"]` 数量回到预期值。
- 没有 DocumentServer 类请求：`/doc/.../c`、`CommandService`、`ConvertService`、`coauthor`、`sockjs`、`websocket`。
- 关闭后 Chrome heap、DOM counters、RSS 不持续随次数增长。

## 真实文件负载测试

启动 Demo 和带 CDP 的 Chrome 后，运行同页多文档真实负载测试：

```bash
npm run test:real-load -- \
  --browser-ws=ws://localhost:9222/devtools/browser \
  --url=http://127.0.0.1:5173/ \
  --file=/path/to/real.docx \
  --file=/path/to/real.pptx \
  --file=/path/to/real.xls
```

默认每轮同时打开 3 个真实文件，并混合覆盖 `edit`、`readonly`、`preview` 三种模式；脚本会等待每个实例触发 `onDocumentReady` 后再停留 2 秒，然后通过 Demo 的 Close All 关闭，循环 60 次并输出 JSON/CSV 到 `test-results/memory/`。如果需要改变并发数，可传 `--batch-size=N`，真实文件模式下需要至少提供 N 个 `--file`。

也可以通过 `--modes=preview` 单独压测上游 embedded 纯预览模式，或通过 `--modes=edit,readonly,preview --batch-size=3` 明确做三模式混合批次。
