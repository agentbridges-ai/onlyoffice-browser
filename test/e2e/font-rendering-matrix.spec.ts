import { existsSync, readdirSync } from 'node:fs';
import { extname, join } from 'node:path';
import { expect, test } from '@playwright/test';

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

test.describe('real Office corpus font rendering', () => {
  test('new Word documents use DengXian as the model and editing default', async ({ page }) => {
    test.setTimeout(60_000);
    await page.goto('/?hardResetOnLastDestroy=true');
    const cacheLaterButton = page.getByRole('button', { name: 'Later' });
    await cacheLaterButton.waitFor({ state: 'visible', timeout: 5_000 }).catch(() => undefined);
    if (await cacheLaterButton.isVisible()) {
      const dengXianRow = page.locator('.font-download-row').filter({ hasText: 'DengXian' });
      const yaHeiRow = page.locator('.font-download-row').filter({ hasText: 'Microsoft YaHei' });
      await expect(dengXianRow.getByRole('button', { name: /Required|Downloaded/ })).toBeDisabled();
      await expect(yaHeiRow.getByRole('button', { name: 'Download' })).toBeEnabled();
      await cacheLaterButton.click();
    }
    await page.locator('.new-menu > summary').click();
    await page.getByRole('button', { name: 'Word document' }).click();
    await page.waitForFunction(
      () => {
        const demo = (window as any).__officeDemo;
        return demo?.tabs?.length === 1 && demo.editor?.getState?.().status === 'ready';
      },
      null,
      { timeout: 25_000 },
    );

    const editorFrame = page
      .frames()
      .find((frame) => /\/documenteditor\/main\/index\.html/.test(frame.url()));
    expect(editorFrame, 'new Word editor frame').toBeTruthy();

    const state = await editorFrame!.evaluate(() => {
      const scope = globalThis as any;
      const defaultTextPr = scope.AscBuilder?.Word?.Api?.GetDocument?.()?.GetDefaultTextPr?.();
      const fontInputs = Array.from(
        document.querySelectorAll<HTMLInputElement>('input[role="combobox"]'),
        (input) => input.value,
      );
      return {
        ascii: defaultTextPr?.GetFontFamily?.('ascii') || '',
        eastAsia: defaultTextPr?.GetFontFamily?.('eastAsia') || '',
        fontInputs,
        callbackPatched: Boolean(scope.Asc?.editor?.__onlyOfficeBrowserFontFamilyCallbackPatched),
      };
    });
    expect(state.ascii).toBe('DengXian');
    expect(state.eastAsia).toBe('DengXian');
    expect(state.callbackPatched).toBe(true);
    expect(state.fontInputs.some((value) => value === 'DengXian' || value === '等线')).toBe(true);
    expect(state.fontInputs).not.toContain('Microsoft YaHei');
    expect(state.fontInputs).not.toContain('微软雅黑');

    await editorFrame!.evaluate(() => {
      const editor = (globalThis as any).Asc.editor;
      editor.sync_TextPrFontFamilyCallBack({ Name: 'Microsoft YaHei' });
      for (let cycle = 0; cycle < 4; cycle += 1) {
        editor.zoomIn();
        editor.zoomOut();
      }
    });
    await page.waitForTimeout(250);

    const fontInputsAfterSelectionAndZoom = await editorFrame!.evaluate(() =>
      Array.from(document.querySelectorAll<HTMLInputElement>('input[role="combobox"]'), (input) => input.value),
    );
    expect(
      fontInputsAfterSelectionAndZoom.some((value) => value === 'DengXian' || value === '等线'),
    ).toBe(true);
    expect(fontInputsAfterSelectionAndZoom).not.toContain('Microsoft YaHei');
    expect(fontInputsAfterSelectionAndZoom).not.toContain('微软雅黑');

    await page.evaluate(() => {
      void (window as any).__officeDemo?.closeAll?.();
    });
  });

  if (matrixFiles.length === 0) {
    test('requires an external matrix directory', async () => {
      test.skip(true, 'Set ONLYOFFICE_MATRIX_DIR to the real Office preview corpus.');
    });
  }

  for (const filePath of matrixFiles) {
    const fileName = filePath.split('/').pop() || filePath;
    test(`${fileName} renders with model-level font substitution`, async ({ page }, testInfo) => {
      test.setTimeout(60_000);
      await page.goto('/?hardResetOnLastDestroy=true');
      const cacheLaterButton = page.getByRole('button', { name: 'Later' });
      if (await cacheLaterButton.isVisible()) await cacheLaterButton.click();
      await page.locator('#file-input').setInputFiles(filePath);
      await page.waitForFunction(
        () => {
          const demo = (window as any).__officeDemo;
          return demo?.tabs?.length === 1 && demo.editor?.getState?.().status === 'ready';
        },
        null,
        { timeout: 25_000 },
      );

      const editorFrame = page
        .frames()
        .find((frame) => /\/(?:document|spreadsheet|presentation)editor\/main\/index\.html/.test(frame.url()));
      expect(editorFrame, `${fileName}: editor frame`).toBeTruthy();

      const rendering = await editorFrame!.evaluate(() => {
        const scope = globalThis as typeof globalThis & {
          __onlyOfficeBrowserFontMetadataFallback?: {
            fallbackFamilyName?: string;
            unavailableFamilyNames?: string[];
            visibleFamilyNames?: string[];
          };
          AscFonts?: {
            g_fontApplication?: {
              GetFontFileWeb?: (name: string, style?: number) => {
                m_wsFontName?: string;
                m_wsFontPath?: string;
                m_lIndex?: number;
              };
              GetFontInfo?: (name: string, style?: number) => { Name?: string };
            };
          };
        };
        const application = scope.AscFonts?.g_fontApplication;
        const fallback = application?.GetFontFileWeb?.('DengXian', 0);
        const unavailable = ['Microsoft YaHei', 'KaiTi', 'SimHei'].map((name) => {
          const selected = application?.GetFontFileWeb?.(name, 0);
          return {
            requested: name,
            selectedName: selected?.m_wsFontName || '',
            selectedPath: selected?.m_wsFontPath || '',
            selectedIndex: selected?.m_lIndex ?? -1,
            infoName: application?.GetFontInfo?.(name, 0)?.Name || '',
          };
        });
        const canvases = Array.from(document.querySelectorAll('canvas'));
        return {
          fallback: {
            name: fallback?.m_wsFontName || '',
            path: fallback?.m_wsFontPath || '',
            index: fallback?.m_lIndex ?? -1,
          },
          unavailable,
          metadataFallback: scope.__onlyOfficeBrowserFontMetadataFallback,
          visibleCanvasCount: canvases.filter((canvas) => canvas.width > 0 && canvas.height > 0).length,
        };
      });

      expect(rendering.fallback.name, `${fileName}: default fallback family`).toBe('DengXian');
      expect(rendering.metadataFallback?.fallbackFamilyName, `${fileName}: pre-init fallback metadata`).toBe(
        'DengXian',
      );
      for (const unavailableName of ['Microsoft YaHei', 'KaiTi', 'SimHei']) {
        expect(
          rendering.metadataFallback?.unavailableFamilyNames,
          `${fileName}: ${unavailableName} remapped before font indexes`,
        ).toContain(unavailableName);
      }
      expect(rendering.visibleCanvasCount, `${fileName}: rendered editor canvases`).toBeGreaterThan(0);
      for (const font of rendering.unavailable) {
        expect(font.selectedName, `${fileName}: ${font.requested} selected family`).toBe('DengXian');
        expect(font.infoName, `${fileName}: ${font.requested} font info`).toBe('DengXian');
        expect(font.selectedPath, `${fileName}: ${font.requested} selected path`).toBe(rendering.fallback.path);
        expect(font.selectedIndex, `${fileName}: ${font.requested} selected face index`).toBe(rendering.fallback.index);
      }

      if (await cacheLaterButton.isVisible()) await cacheLaterButton.click();
      const zoomKind = await editorFrame!.evaluate(() => {
        const editor = (globalThis as any).Asc?.editor;
        if (typeof editor?.zoomIn === 'function' && typeof editor?.zoomOut === 'function') return 'step';
        if (typeof editor?.asc_getZoom === 'function' && typeof editor?.asc_setZoom === 'function') return 'scale';
        return '';
      });
      expect(zoomKind, `${fileName}: zoom API`).not.toBe('');
      for (let cycle = 0; cycle < 4; cycle += 1) {
        await editorFrame!.evaluate((kind) => {
          const editor = (globalThis as any).Asc.editor;
          if (kind === 'step') {
            editor.zoomIn();
            editor.zoomOut();
          } else {
            const zoom = editor.asc_getZoom();
            editor.asc_setZoom(Math.min(zoom + 0.25, 5));
            editor.asc_setZoom(zoom);
          }
        }, zoomKind);
        await page.waitForTimeout(75);
      }
      await page.waitForTimeout(250);

      const renderingAfterZoom = await editorFrame!.evaluate(() => {
        const scope = globalThis as typeof globalThis & {
          AscFonts?: {
            g_fontApplication?: {
              GetFontFileWeb?: (name: string, style?: number) => {
                m_wsFontName?: string;
                m_wsFontPath?: string;
                m_lIndex?: number;
              };
            };
          };
        };
        const application = scope.AscFonts?.g_fontApplication;
        return ['Microsoft YaHei', 'KaiTi', 'SimHei'].map((name) => {
          const selected = application?.GetFontFileWeb?.(name, 0);
          return {
            requested: name,
            selectedName: selected?.m_wsFontName || '',
            selectedPath: selected?.m_wsFontPath || '',
            selectedIndex: selected?.m_lIndex ?? -1,
          };
        });
      });
      for (const font of renderingAfterZoom) {
        expect(font.selectedName, `${fileName}: ${font.requested} family after zoom`).toBe('DengXian');
        expect(font.selectedPath, `${fileName}: ${font.requested} path after zoom`).toBe(rendering.fallback.path);
        expect(font.selectedIndex, `${fileName}: ${font.requested} face after zoom`).toBe(rendering.fallback.index);
      }

      if (filePath === realDocx) {
        if (await cacheLaterButton.isVisible()) await cacheLaterButton.click();
        await page.screenshot({
          path: testInfo.outputPath('real-docx-font-rendering-after-zoom-cycles.png'),
          fullPage: false,
        });
      }

      await page.evaluate(() => {
        void (window as any).__officeDemo?.closeAll?.();
      });
      await page.waitForFunction(
        () =>
          ((window as any).__officeDemo?.tabs?.length || 0) === 0 &&
          !(window as any).__officeDemo?.editor &&
          document.querySelectorAll('iframe.office-editor-host-frame, iframe[name="frameEditor"]').length === 0,
      );
    });
  }
});
