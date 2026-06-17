import path, { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
import type { Plugin } from 'vite';

function serveWorkspaceSrc(): Plugin {
  const sourceRoot = resolve(fileURLToPath(new URL('.', import.meta.url)), 'src');
  const toFileSystemUrl = (filePath: string): string => `/@fs${filePath}`;

  return {
    name: 'serve-workspace-src',
    enforce: 'pre',
    transformIndexHtml: {
      order: 'pre',
      handler(html, ctx) {
        if (!ctx.server) return html;

        return html.replace(
          /(<script\s+type="module"\s+src=")(?:\.\.\/)+src\/index\.ts("><\/script>)/g,
          `$1${toFileSystemUrl(path.join(sourceRoot, 'index.ts'))}$2`,
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
  plugins: [serveWorkspaceSrc()],
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
