import fs from 'node:fs';
import type { IncomingMessage, ServerResponse } from 'node:http';
import path, { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
import type { Plugin } from 'vite';

export const FONT_ASSETS_DIR_ENV = 'ONLYOFFICE_BROWSER_FONT_ASSETS_DIR';
export const GENERATED_FONT_ASSETS_MANIFEST = 'onlyoffice-browser-font-assets.json';
export const GENERATED_FONT_SOURCE_MAP = 'onlyoffice-browser-font-source-map.json';

type NextFunction = (error?: unknown) => void;

const FONT_THUMBNAIL_RE = /^\/sdkjs\/common\/Images\/fonts_thumbnail(?:_ea)?(?:@[\d.]+x)?\.png$/;
const FONT_ASSET_MIME_TYPES: Record<string, string> = {
  '.bin': 'application/octet-stream',
  '.eot': 'application/vnd.ms-fontobject',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.otc': 'font/collection',
  '.otf': 'font/otf',
  '.png': 'image/png',
  '.ttc': 'font/collection',
  '.tte': 'font/ttf',
  '.ttf': 'font/ttf',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

function isInsideDirectory(root: string, candidate: string): boolean {
  const relative = path.relative(root, candidate);
  return !!relative && !relative.startsWith('..') && !path.isAbsolute(relative);
}

function safeResolve(root: string, relativePath: string): string | null {
  const resolved = path.resolve(root, relativePath);
  return isInsideDirectory(root, resolved) ? resolved : null;
}

export function resolveGeneratedFontAssetPath(fontAssetsRoot: string, requestUrl: string): string | null {
  let pathname: string;
  try {
    pathname = new URL(requestUrl, 'http://onlyoffice-browser.local').pathname;
  } catch {
    return null;
  }

  const normalizedPathname = pathname.replace(/\/{2,}/g, '/');

  if (normalizedPathname === `/${GENERATED_FONT_ASSETS_MANIFEST}`) {
    return safeResolve(fontAssetsRoot, GENERATED_FONT_ASSETS_MANIFEST);
  }

  if (normalizedPathname === `/${GENERATED_FONT_SOURCE_MAP}`) {
    return safeResolve(fontAssetsRoot, GENERATED_FONT_SOURCE_MAP);
  }

  if (normalizedPathname === '/sdkjs/common/AllFonts.js') {
    return safeResolve(fontAssetsRoot, 'sdkjs/common/AllFonts.js');
  }

  if (FONT_THUMBNAIL_RE.test(normalizedPathname)) {
    return safeResolve(fontAssetsRoot, `sdkjs/common/Images/${path.basename(normalizedPathname)}`);
  }

  if (normalizedPathname === '/server/FileConverter/bin/font_selection.bin') {
    return safeResolve(fontAssetsRoot, 'server/FileConverter/bin/font_selection.bin');
  }

  if (normalizedPathname === '/server/FileConverter/bin/AllFonts.js') {
    return safeResolve(fontAssetsRoot, 'server/FileConverter/bin/AllFonts.js');
  }

  if (normalizedPathname.startsWith('/fonts/fonts/')) {
    return null;
  }

  if (normalizedPathname.startsWith('/fonts/')) {
    return safeResolve(fontAssetsRoot, `fonts/${normalizedPathname.slice('/fonts/'.length)}`);
  }

  return null;
}

function serveStaticFile(filePath: string, res: ServerResponse): void {
  const contentType = FONT_ASSET_MIME_TYPES[path.extname(filePath).toLowerCase()] || 'application/octet-stream';
  res.statusCode = 200;
  res.setHeader('Content-Type', contentType);
  res.setHeader('Cache-Control', 'no-cache');
  fs.createReadStream(filePath).pipe(res);
}

function serveGeneratedFontAssets(): Plugin {
  const configuredDir = process.env[FONT_ASSETS_DIR_ENV]?.trim();
  const fontAssetsRoot = configuredDir ? path.resolve(configuredDir) : '';
  let hasWarnedMissingDirectory = false;

  const middleware = (req: IncomingMessage, res: ServerResponse, next: NextFunction): void => {
    if (!configuredDir || !req.url) {
      next();
      return;
    }

    if (!fs.existsSync(fontAssetsRoot) || !fs.statSync(fontAssetsRoot).isDirectory()) {
      if (!hasWarnedMissingDirectory) {
        console.warn(`${FONT_ASSETS_DIR_ENV} points to a missing directory: ${fontAssetsRoot}`);
        hasWarnedMissingDirectory = true;
      }
      next();
      return;
    }

    const assetPath = resolveGeneratedFontAssetPath(fontAssetsRoot, req.url);
    if (!assetPath || !fs.existsSync(assetPath) || !fs.statSync(assetPath).isFile()) {
      next();
      return;
    }

    serveStaticFile(assetPath, res);
  };

  return {
    name: 'serve-generated-font-assets',
    enforce: 'pre',
    configureServer(server) {
      server.middlewares.use(middleware);
    },
    configurePreviewServer(server) {
      server.middlewares.use(middleware);
    },
  };
}

function serveWorkspaceSrc(): Plugin {
  const sourceRoot = resolve(__dirname, 'src');
  const toFileSystemUrl = (filePath: string): string => `/@fs${filePath}`;

  return {
    name: 'serve-workspace-src',
    enforce: 'pre',
    transformIndexHtml: {
      order: 'pre',
      handler(html, ctx) {
        if (!ctx.server) return html;

        return html.replace(
          /(<script\s+type="module"\s+src=")(?:\.\.\/)+src\/([^"]+\.ts)("><\/script>)/g,
          (_match, prefix: string, sourceFile: string, suffix: string) =>
            `${prefix}${toFileSystemUrl(path.join(sourceRoot, sourceFile))}${suffix}`,
        );
      },
    },
    configureServer(server) {
      server.middlewares.use((req, _res, next) => {
        if (req.url?.startsWith('/src/')) {
          const [requestPath, query = ''] = req.url.slice('/src/'.length).split('?', 2);
          req.url = `${toFileSystemUrl(path.join(sourceRoot, requestPath))}${query ? `?${query}` : ''}`;
        }
        next();
      });
    },
  };
}

const __filename = fileURLToPath(import.meta.url);

const __dirname = path.dirname(__filename);

export default defineConfig({
  root: 'pages',
  base: './',
  publicDir: resolve(__dirname, 'public'),
  plugins: [serveGeneratedFontAssets(), serveWorkspaceSrc()],
  server: {
    fs: {
      // Allow Vite to serve src/ which lives outside the pages/ root
      allow: [__dirname],
    },
  },
  build: {
    outDir: resolve(__dirname, 'dist'),
    emptyOutDir: true,
    rollupOptions: {
      input: {
        main: resolve(__dirname, 'pages/index.html'),
        officeHost: resolve(__dirname, 'pages/office-host.html'),
        saveE2E: resolve(__dirname, 'pages/save-e2e.html'),
      },
    },
  },
  resolve: {
    alias: {
      '@/lib': resolve(__dirname, 'src/lib'),
      '@/store': resolve(__dirname, 'src/store'),
      '@/types': resolve(__dirname, 'src/types'),
      '@/styles': resolve(__dirname, 'src/styles'),
    },
    extensions: ['.mjs', '.js', '.ts', '.jsx', '.tsx', '.json'],
  },
  css: {
    preprocessorOptions: {
      scss: {
        additionalData: `@import "@/styles/base.css";`,
      },
    },
  },
});
