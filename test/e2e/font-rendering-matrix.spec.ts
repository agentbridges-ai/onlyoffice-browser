import { existsSync, readdirSync } from 'node:fs';
import { extname, join } from 'node:path';
import { expect, test, type Frame, type Page } from '@playwright/test';

const matrixDirectory = process.env.ONLYOFFICE_MATRIX_DIR;
const realDocx = process.env.ONLYOFFICE_REAL_DOCX;
const supportedExtensions = new Set([
  '.csv',
  '.doc',
  '.docx',
  '.odp',
  '.ods',
  '.odt',
  '.ppt',
  '.pptx',
  '.rtf',
  '.xls',
  '.xlsx',
]);

const matrixFiles =
  matrixDirectory && existsSync(matrixDirectory)
    ? readdirSync(matrixDirectory)
        .filter((name) => supportedExtensions.has(extname(name).toLowerCase()))
        .sort()
        .map((name) => join(matrixDirectory, name))
    : [];
if (realDocx && existsSync(realDocx)) matrixFiles.push(realDocx);

async function dismissLegacyResourcePrompt(page: Page): Promise<void> {
  const later = page.getByRole('button', { name: 'Later' });
  await later.waitFor({ state: 'visible', timeout: 5_000 }).catch(() => undefined);
  if (await later.isVisible()) await later.click();
}

async function waitForEditor(page: Page): Promise<Frame> {
  await page.waitForFunction(
    () => {
      const demo = (window as any).__officeDemo;
      return demo?.tabs?.length === 1 && demo.editor?.getState?.().status === 'ready';
    },
    null,
    { timeout: 25_000 },
  );
  const frame = page
    .frames()
    .find((candidate) => /\/(?:document|spreadsheet|presentation)editor\/main\/index\.html/.test(candidate.url()));
  expect(frame, 'OnlyOffice editor frame').toBeTruthy();
  return frame!;
}

test.describe('OnlyOffice native font behavior', () => {
  test('new Word documents use Aptos for Western text and DengXian for East Asian text', async ({ page }) => {
    test.setTimeout(60_000);
    await page.goto('/?hardResetOnLastDestroy=true');
    await dismissLegacyResourcePrompt(page);
    await page.locator('.new-menu > summary').click();
    await page.getByRole('button', { name: 'Word document' }).click();
    const editorFrame = await waitForEditor(page);

    const defaults = await editorFrame.evaluate(() => {
      const scope = globalThis as any;
      const textPr = scope.AscBuilder?.Word?.Api?.GetDocument?.()?.GetDefaultTextPr?.();
      return {
        ascii: textPr?.GetFontFamily?.('ascii') || '',
        hAnsi: textPr?.GetFontFamily?.('hAnsi') || '',
        eastAsia: textPr?.GetFontFamily?.('eastAsia') || '',
        forcedFileFallback: Boolean(scope.AscFonts?.g_fontApplication?.__onlyOfficeBrowserFontFileFallback),
        forcedCallbackFallback: Boolean(scope.Asc?.editor?.__onlyOfficeBrowserFontFamilyCallbackPatched),
      };
    });

    expect(defaults).toEqual({
      ascii: 'Aptos',
      hAnsi: 'Aptos',
      eastAsia: 'DengXian',
      forcedFileFallback: false,
      forcedCallbackFallback: false,
    });
  });

  test('every packaged font family resolves to its own installed face and can be applied', async ({ page }) => {
    test.setTimeout(3 * 60_000);
    await page.goto('/?hardResetOnLastDestroy=true');
    await dismissLegacyResourcePrompt(page);
    const familyNames = await page.evaluate(async () => {
      const response = await fetch('/onlyoffice-browser-font-assets.json', { cache: 'no-store' });
      const manifest = await response.json();
      return (manifest.fontFamilies || []).map((family: { name: string }) => family.name);
    });
    expect(
      familyNames.length,
      'the complete Office package must expose at least 170 font families',
    ).toBeGreaterThanOrEqual(170);
    expect(familyNames).toContain('Aptos');
    expect(familyNames).toContain('Calibri');
    expect(familyNames).toContain('DengXian');

    await page.locator('.new-menu > summary').click();
    await page.getByRole('button', { name: 'Word document' }).click();
    const editorFrame = await waitForEditor(page);
    const results = await editorFrame.evaluate((names) => {
      const scope = globalThis as any;
      const application = scope.AscFonts?.g_fontApplication;
      const editor = scope.Asc?.editor;
      return names.map((name: string) => {
        editor?.put_TextPrFontName?.(name);
        const styles = [0, 1, 2, 3].map((style) => {
          const selected = application?.GetFontFileWeb?.(name, style);
          return {
            style,
            selectedName: selected?.m_wsFontName || '',
            selectedPath: selected?.m_wsFontPath || '',
            selectedIndex: selected?.m_lIndex ?? -1,
          };
        });
        return { name, styles };
      });
    }, familyNames);

    for (const family of results) {
      for (const style of family.styles) {
        expect(style.selectedName, `${family.name} style ${style.style} selected family`).toBe(family.name);
        expect(style.selectedPath, `${family.name} style ${style.style} font path`).not.toBe('');
        expect(style.selectedIndex, `${family.name} style ${style.style} face index`).toBeGreaterThanOrEqual(0);
      }
    }
  });

  test('missing fonts use DocumentServer native closest-match selection and remain stable', async ({ page }) => {
    test.setTimeout(60_000);
    await page.goto('/?hardResetOnLastDestroy=true');
    await dismissLegacyResourcePrompt(page);
    await page.locator('.new-menu > summary').click();
    await page.getByRole('button', { name: 'Word document' }).click();
    const editorFrame = await waitForEditor(page);
    const missingNames = ['OnlyOffice Missing Sans', 'OnlyOffice Missing Serif', '不存在的字体'];

    const select = () =>
      editorFrame.evaluate((names) => {
        const scope = globalThis as any;
        const application = scope.AscFonts?.g_fontApplication;
        return {
          forcedFileFallback: Boolean(application?.__onlyOfficeBrowserFontFileFallback),
          forcedCallbackFallback: Boolean(scope.Asc?.editor?.__onlyOfficeBrowserFontFamilyCallbackPatched),
          selected: names.map((name) => {
            const font = application?.GetFontFileWeb?.(name, 0);
            return {
              requested: name,
              name: font?.m_wsFontName || '',
              path: font?.m_wsFontPath || '',
              index: font?.m_lIndex ?? -1,
            };
          }),
        };
      }, missingNames);

    const before = await select();
    expect(before.forcedFileFallback).toBe(false);
    expect(before.forcedCallbackFallback).toBe(false);
    for (const font of before.selected) {
      expect(font.name, `${font.requested} native substitute`).not.toBe('');
      expect(font.path, `${font.requested} native substitute path`).not.toBe('');
      expect(font.index, `${font.requested} native substitute face`).toBeGreaterThanOrEqual(0);
    }

    await editorFrame.evaluate(() => {
      const editor = (globalThis as any).Asc?.editor;
      for (let cycle = 0; cycle < 4; cycle += 1) {
        editor?.zoomIn?.();
        editor?.zoomOut?.();
      }
    });
    expect(await select()).toEqual(before);
  });

  if (matrixFiles.length === 0) {
    test('requires an external matrix directory', async () => {
      test.skip(true, 'Set ONLYOFFICE_MATRIX_DIR to the real Office preview corpus.');
    });
  }

  for (const filePath of matrixFiles) {
    const fileName = filePath.split('/').pop() || filePath;
    test(`${fileName} renders through the native font dictionary across zoom cycles`, async ({ page }, testInfo) => {
      test.setTimeout(60_000);
      await page.goto('/?hardResetOnLastDestroy=true');
      await dismissLegacyResourcePrompt(page);
      await page.locator('#file-input').setInputFiles(filePath);
      const editorFrame = await waitForEditor(page);
      const before = await editorFrame.evaluate(() => ({
        canvases: Array.from(document.querySelectorAll('canvas')).filter(
          (canvas) => canvas.width > 0 && canvas.height > 0,
        ).length,
        forcedFileFallback: Boolean(
          (globalThis as any).AscFonts?.g_fontApplication?.__onlyOfficeBrowserFontFileFallback,
        ),
      }));
      expect(before.canvases, `${fileName}: rendered editor canvases`).toBeGreaterThan(0);
      expect(before.forcedFileFallback, `${fileName}: native fallback path`).toBe(false);

      await editorFrame.evaluate(() => {
        const editor = (globalThis as any).Asc?.editor;
        for (let cycle = 0; cycle < 4; cycle += 1) {
          if (typeof editor?.zoomIn === 'function') {
            editor.zoomIn();
            editor.zoomOut();
          } else {
            const zoom = editor.asc_getZoom();
            editor.asc_setZoom(Math.min(zoom + 0.25, 5));
            editor.asc_setZoom(zoom);
          }
        }
      });
      await page.waitForTimeout(250);
      expect(
        await editorFrame.locator('canvas').evaluateAll(
          (canvases) =>
            canvases.filter((canvas) => {
              const rendered = canvas as HTMLCanvasElement;
              return rendered.width > 0 && rendered.height > 0;
            }).length,
        ),
      ).toBeGreaterThan(0);

      if (filePath === realDocx) {
        await page.screenshot({
          path: testInfo.outputPath('real-docx-native-font-rendering-after-zoom-cycles.png'),
          fullPage: false,
        });
      }
    });
  }
});
