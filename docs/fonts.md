# Font Management

## Default behavior

This package does not ship any browser runtime font files, generated font registries, or font picker thumbnails. Font assets are mandatory: the editor host must serve a generated OnlyOffice font asset directory before any document is opened.

OnlyOffice's official installation flow adds fonts by copying them into the DocumentServer font directory and running `documentserver-generate-allfonts.sh`. That generator updates the font registry, font picker thumbnails, web font files, and converter font selection data.

Official references: [Docker](https://helpcenter.onlyoffice.com/docs/installation/docs-install-fonts-docker.aspx), [Linux](https://helpcenter.onlyoffice.com/docs/installation/docs-community-install-fonts-linux.aspx), [Windows](https://helpcenter.onlyoffice.com/docs/installation/docs-community-install-fonts-windows.aspx).

## Generate font assets

Use Docker to run the official OnlyOffice generator against a directory of fonts you are allowed to use:

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets
```

Installed package users can run the same generator through the published CLI:

```bash
npx onlyoffice-browser-generate-font-assets \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets
```

By default the script exports a compact `zh-core` set focused on Simplified Chinese Office documents and classic English fonts. The visible font registry is limited to Microsoft YaHei, Microsoft YaHei UI, SimSun, NSimSun, SimSun-ExtB, SimHei, KaiTi, FangSong, DengXian, Aptos, Calibri, Arial, Times New Roman, Consolas, Cambria, and Cambria Math. It keeps these source files only: Microsoft YaHei regular/bold, SimSun/NSimSun, SimSun-ExtB, SimHei, KaiTi, FangSong, DengXian regular/bold, Aptos regular/italic/bold/bold-italic, Calibri regular/italic/bold/bold-italic, Arial regular/italic/bold/bold-italic, Times New Roman regular/italic/bold/bold-italic, Consolas regular/bold, and Cambria/Cambria Math regular. Traditional Chinese and other language-specific fonts are omitted by default to keep first paint fast.

Use the full official generated set only when you need broad language coverage:

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets \
  --font-set full
```

Add an exact family to the `zh-core` set when a document requires it:

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets \
  --keep-font Wingdings
```

The default generator image is `onlyoffice/documentserver:9.3.0`. Override it when needed:

```bash
npm run fonts:generate -- \
  --input /absolute/path/to/fonts \
  --output .onlyoffice-font-assets \
  --image onlyoffice/documentserver:9.3.0
```

The output directory contains the generated OnlyOffice assets:

- `sdkjs/common/AllFonts.js`
- `sdkjs/common/Images/fonts_thumbnail*.png`
- `sdkjs/common/Images/fonts_thumbnail_ea*.png`
- `fonts/`
- `server/FileConverter/bin/font_selection.bin`
- `onlyoffice-browser-font-assets.json`

The script keeps the official generated registry, thumbnails, and converter metadata, then rewrites the web font file list to point at raw font copies under the canonical `/fonts/<file>` mount path. This matches the current browser runtime, which loads original TTF/OTF/TTC/WOFF files instead of DocumentServer's internal extensionless web font binaries.

Verify the output before starting the editor:

```bash
npm run fonts:verify -- --input .onlyoffice-font-assets
```

Installed package users can use:

```bash
npx onlyoffice-browser-verify-font-assets --input .onlyoffice-font-assets
```

## Use generated assets

For local development and visual QA, start the demo with:

```bash
npm run dev:fonts
```

The dev and preview servers serve files from that generated directory at the same paths production must expose. There is no default font fallback in the package. If `onlyoffice-browser-font-assets.json`, `AllFonts.js`, thumbnails, `/fonts/<file>`, or `font_selection.bin` are missing, editor startup fails with a configuration error.

The editor receives the generated `AllFonts.js`, thumbnail images, font files, and converter metadata as static assets. For compact sets such as `zh-core`, the full generated font registry remains available for document font resolution, while the editor UI font picker is filtered to `__fonts_visible_names`.

For production, deploy the generated directory on the same editor host path layout. The manifest lets the browser-side converter load the same font files into x2t's `/working/fonts/` directory and load `font_selection.bin` for conversion.

## Licensing

The generator accepts open-source fonts, internally licensed fonts, vendor-provided fonts, or system font copies. You are responsible for ensuring that you have the right to use and deploy those files.

Do not commit generated third-party font assets to this repository unless their license explicitly allows redistribution.
