# 字体管理

## 默认行为

本包不内置任何 browser runtime 字体文件、生成后的字体注册表或字体下拉缩略图。字体资产是必需配置：editor host 必须先提供一套生成后的 OnlyOffice 字体资产目录，才能打开文档。

OnlyOffice 官方安装流程是在 DocumentServer 字体目录中加入字体，然后运行 `documentserver-generate-allfonts.sh`。该生成器会更新字体注册表、字体下拉缩略图、Web 字体文件和转换器字体选择数据。

官方参考：[Docker](https://helpcenter.onlyoffice.com/docs/installation/docs-install-fonts-docker.aspx)、[Linux](https://helpcenter.onlyoffice.com/docs/installation/docs-community-install-fonts-linux.aspx)、[Windows](https://helpcenter.onlyoffice.com/docs/installation/docs-community-install-fonts-windows.aspx)。

## 生成字体资产

使用 Docker 运行官方 OnlyOffice 生成器，输入目录放置你有权使用的字体文件：

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets
```

如果是在安装包的项目中使用，可以直接运行发布出来的 CLI：

```bash
npx onlyoffice-browser-generate-font-assets \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets
```

脚本默认导出精简的 `zh-core` 字体集，集中在简体中文 Office 文档和经典英文字体。可见字体注册表只保留微软雅黑、Microsoft YaHei UI、宋体、新宋体、SimSun-ExtB、黑体、楷体、仿宋、等线、Aptos、Calibri、Arial、Times New Roman、Consolas、Cambria 和 Cambria Math。默认只保留这些字体源文件：微软雅黑 regular/bold、宋体/新宋体、SimSun-ExtB、黑体、楷体、仿宋、等线 regular/bold、Aptos regular/italic/bold/bold-italic、Calibri regular/italic/bold/bold-italic、Arial regular/italic/bold/bold-italic、Times New Roman regular/italic/bold/bold-italic、Consolas regular/bold、Cambria/Cambria Math regular。繁体中文和其他语言专属字体默认不导出，以减少首屏字体加载。

确实需要广泛语言覆盖时，可以导出完整官方生成集：

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets \
  --font-set full
```

如果某个文档需要 `zh-core` 集之外的具体字体族，可以追加精确字体名：

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets \
  --keep-font Wingdings
```

默认镜像是 `onlyoffice/documentserver:9.3.0`。需要时可以覆盖：

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets \
  --image onlyoffice/documentserver:9.3.0
```

输出目录包含这些 OnlyOffice 生成资产：

- `sdkjs/common/AllFonts.js`
- `sdkjs/common/Images/fonts_thumbnail*.png`
- `sdkjs/common/Images/fonts_thumbnail_ea*.png`
- `fonts/`
- `server/FileConverter/bin/font_selection.bin`
- `onlyoffice-browser-font-assets.json`

脚本会保留官方生成的字体注册表、缩略图和转换器元数据，然后把 Web 字体文件列表改写为指向统一挂载路径 `/fonts/<file>` 下的原始字体副本。这与当前 browser runtime 匹配；当前 runtime 加载原始 TTF/OTF/TTC/WOFF 文件，而不是 DocumentServer 内部使用的无扩展名 web font binary。

启动编辑器前先验证输出目录：

```bash
npm run fonts:verify -- --input .onlyoffice-font-assets
```

安装包用户也可以使用：

```bash
npx onlyoffice-browser-verify-font-assets --input .onlyoffice-font-assets
```

## 使用生成资产

本地开发和视觉验收时这样启动 demo：

```bash
npm run dev:fonts
```

dev 和 preview server 会把生成目录按生产同款路径暴露出来。本包没有默认字体 fallback；如果缺少 `onlyoffice-browser-font-assets.json`、`AllFonts.js`、字体缩略图、`/fonts/<file>` 或 `font_selection.bin`，编辑器启动会直接报配置错误。

编辑器拿到的是静态生成出的 `AllFonts.js`、字体下拉缩略图、字体文件和转换器元数据。对于 `zh-core` 这类精简集，完整字体注册表仍保留给文档字体解析使用，编辑器 UI 的字体下拉会按 `__fonts_visible_names` 过滤显示。

生产环境应按相同路径结构部署生成目录。manifest 会让浏览器端转换器把同一批字体文件装载进 x2t 的 `/working/fonts/`，并加载 `font_selection.bin`，保证转换侧和编辑器侧尽量一致。

## 授权

生成器可以接收开源字体、内部授权字体、厂商提供字体或系统字体副本。是否有权使用和部署这些字体文件，由使用者自行负责。

除非字体许可证明确允许再分发，否则不要把生成出的第三方字体资产提交到本仓库。
