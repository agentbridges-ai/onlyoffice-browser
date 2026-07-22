import { expect, type Frame, type Page, test } from '@playwright/test';

const FIXTURE_URL = '/fixtures/regressions/example-document-title-ole.doc';
const FIXTURE_NAME = 'Example Document Title.doc';
const FIXTURE_SHA256 = 'd85e44ae5368ccbbe57ded8533ced05a250c30cfa15da10f19fdaf63f080238c';
const ODT_FIXTURE_URL = '/fixtures/regressions/example-document-title.odt';

type SaveE2EStatus = {
  type: string;
  ready: boolean;
  initialHash: string;
  initialSize: number;
  error: string;
  state: { status: string } | null;
};

// Each page instantiates the full x2t-wasm/editor stack. Keep these large
// regression documents independent without running two converters in the same
// browser worker at once; the per-test timeout remains the project default 30s.
test.describe.configure({ mode: 'default' });

test('legacy DOC decodes its embedded chart preview instead of only reaching READY', async ({ page }) => {
  const failures = collectPageFailures(page);
  const params = new URLSearchParams({
    scenario: 'local-file',
    type: 'doc',
    fixtureUrl: FIXTURE_URL,
    fixtureName: FIXTURE_NAME,
  });

  await page.goto(`/save-e2e.html?${params}`);
  await page.waitForFunction(
    () => {
      const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
      return status?.ready === true || Boolean(status?.error);
    },
    null,
    { timeout: 25_000 },
  );

  const status = await getStatus(page);
  expect(status.error).toBe('');
  expect(status.type).toBe('doc');
  expect(status.ready).toBe(true);
  expect(status.state?.status).toBe('ready');
  expect(status.initialSize).toBeGreaterThan(0);
  expect(status.initialHash).toBe(FIXTURE_SHA256);

  const editorFrame = await getWordEditorFrame(page);
  await editorFrame.waitForFunction(() => {
    const scope = globalThis as typeof globalThis & {
      AscCommon?: {
        g_oDocumentUrls?: { urls?: Record<string, string> };
        g_image_loader?: { map_image_index?: Record<string, { Image?: HTMLImageElement }> };
      };
      editor?: {
        WordControl?: {
          m_oLogicDocument?: {
            DrawingObjects?: {
              drawingObjects?: Array<{ GraphicObj?: { isOleObject?: () => boolean; getImageUrl?: () => string } }>;
            };
          };
        };
        ImageLoader?: { map_image_index?: Record<string, { Image?: HTMLImageElement }> };
      };
    };
    const drawings = scope.editor?.WordControl?.m_oLogicDocument?.DrawingObjects?.drawingObjects || [];
    const ole = drawings.find((drawing) => drawing.GraphicObj?.isOleObject?.());
    const imageName = ole?.GraphicObj?.getImageUrl?.();
    if (!imageName) return false;
    const resourceUrl = scope.AscCommon?.g_oDocumentUrls?.urls?.[`media/${imageName}`];
    const image = resourceUrl
      ? (scope.editor?.ImageLoader?.map_image_index?.[resourceUrl] ||
          scope.AscCommon?.g_image_loader?.map_image_index?.[resourceUrl])?.Image
      : undefined;
    return Boolean(image?.complete && image.naturalWidth > 0 && image.naturalHeight > 0);
  });

  const olePreview = await editorFrame.evaluate(() => {
    const scope = globalThis as typeof globalThis & {
      AscCommon: { g_oDocumentUrls: { urls: Record<string, string> } };
      editor: {
        WordControl: {
          m_oLogicDocument: {
            DrawingObjects: {
              drawingObjects: Array<{ GraphicObj?: { isOleObject?: () => boolean; getImageUrl?: () => string } }>;
            };
          };
        };
        ImageLoader: { map_image_index: Record<string, { Image: HTMLImageElement }> };
      };
    };
    const drawings = scope.editor.WordControl.m_oLogicDocument.DrawingObjects.drawingObjects;
    const ole = drawings.find((drawing) => drawing.GraphicObj?.isOleObject?.());
    const imageName = ole?.GraphicObj?.getImageUrl?.() || '';
    const resourceUrl = scope.AscCommon.g_oDocumentUrls.urls[`media/${imageName}`];
    const image = scope.editor.ImageLoader.map_image_index[resourceUrl]?.Image;
    return {
      imageName,
      naturalWidth: image?.naturalWidth || 0,
      naturalHeight: image?.naturalHeight || 0,
    };
  });

  expect(olePreview.imageName).toBe('display6image1.svg');
  expect(olePreview.naturalWidth).toBeGreaterThan(0);
  expect(olePreview.naturalHeight).toBeGreaterThan(0);
  expect(failures).toEqual([]);
});

test('ODT chart retains five distinct gradient series in the editor document model', async ({ page }) => {
  const failures = collectPageFailures(page);
  const params = new URLSearchParams({
    scenario: 'local-file',
    type: 'odt',
    fixtureUrl: ODT_FIXTURE_URL,
    fixtureName: 'Example Document Title.odt',
  });

  await page.goto(`/save-e2e.html?${params}`);
  await page.waitForFunction(
    () => {
      const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
      return status?.ready === true || Boolean(status?.error);
    },
    null,
    { timeout: 25_000 },
  );
  const status = await getStatus(page);
  expect(status.error).toBe('');
  expect(status.ready).toBe(true);

  const editorFrame = await getWordEditorFrame(page);
  const series = await editorFrame.evaluate(() => {
    type Rgba = { R: number; G: number; B: number; A: number };
    type GradientStop = { color?: { RGBA?: Rgba; color?: { RGBA?: Rgba } } };
    type Series = {
      spPr?: { Fill?: { fill?: { constructor?: { name?: string }; colors?: GradientStop[] } } };
    };
    const scope = globalThis as typeof globalThis & {
      editor: {
        WordControl: {
          m_oLogicDocument: {
            DrawingObjects: {
              drawingObjects: Array<{
                GraphicObj?: {
                  isChart?: () => boolean;
                  chart?: { plotArea?: { charts?: Array<{ series?: Series[] }> } };
                };
              }>;
            };
          };
        };
      };
    };
    const drawings = scope.editor.WordControl.m_oLogicDocument.DrawingObjects.drawingObjects;
    const chart = drawings.find((drawing) => drawing.GraphicObj?.isChart?.())?.GraphicObj?.chart;
    return (chart?.plotArea?.charts?.[0]?.series || []).map((item) => {
      const fill = item.spPr?.Fill?.fill;
      return {
        fillType: fill?.constructor?.name || '',
        colors: (fill?.colors || [])
          // CUniColor.RGBA is a calculated cache and can still be zero before
          // the chart is painted. The nested base color is the serialized
          // sRGB value that drives that calculation.
          .map((stop) => stop.color?.color?.RGBA || stop.color?.RGBA)
          .filter((color): color is Rgba => Boolean(color)),
      };
    });
  });

  expect(series).toHaveLength(5);
  expect(series.map((item) => item.fillType)).toEqual(Array(5).fill('CGradFill'));
  const gradientColors = series.map((item) =>
    item.colors.map(({ R, G, B }) => [R, G, B]),
  );
  expect(gradientColors).toEqual([
    [
      [0x35, 0x74, 0xac],
      [0x46, 0x97, 0xe0],
      [0x43, 0x97, 0xe4],
    ],
    [
      [0xc5, 0x59, 0x0f],
      [0xff, 0x74, 0x15],
      [0xff, 0x74, 0x16],
    ],
    [
      [0x7b, 0x7b, 0x7b],
      [0x9f, 0x9f, 0x9f],
      [0xa0, 0xa0, 0xa0],
    ],
    [
      [0xbe, 0x8f, 0x00],
      [0xf7, 0xba, 0x00],
      [0xf8, 0xba, 0x00],
    ],
    [
      [0x24, 0x51, 0xa0],
      [0x2e, 0x69, 0xd0],
      [0x2c, 0x68, 0xd4],
    ],
  ]);
  expect(failures).toEqual([]);
});

test('DOCX source retains readable body text and five chart series in the editor model', async ({ page }) => {
  const failures = collectPageFailures(page);
  const params = new URLSearchParams({
    scenario: 'local-file',
    type: 'docx',
    fixtureUrl: '/fixtures/regressions/example-document-title.docx',
    fixtureName: 'Example Document Title.docx',
  });
  await page.goto(`/save-e2e.html?${params}`);
  await page.waitForFunction(
    () => {
      const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
      return status?.ready === true || Boolean(status?.error);
    },
    null,
    { timeout: 25_000 },
  );
  expect((await getStatus(page)).error).toBe('');
  const editorFrame = await getWordEditorFrame(page);
  const model = await editorFrame.evaluate(() => {
    const scope = globalThis as typeof globalThis & { editor: any };
    const document = scope.editor.WordControl.m_oLogicDocument;
    const chart = (document.DrawingObjects?.drawingObjects || [])
      .map((drawing: any) => drawing.GraphicObj)
      .find((graphic: any) => graphic?.isChart?.());
    const series = chart?.chart?.plotArea?.charts?.[0]?.series || [];
    const drawingText = (document.DrawingObjects?.drawingObjects || [])
      .flatMap((drawing: any) => {
        const graphic = drawing.GraphicObj;
        const paragraphs =
          graphic?.txBody?.content?.GetAllParagraphs?.() || graphic?.textBoxContent?.GetAllParagraphs?.() || [];
        return paragraphs.map((paragraph: any) => paragraph.GetText?.() || '');
      })
      .join('\n');
    return {
      text: `${(document.GetAllParagraphs?.() || []).map((paragraph: any) => paragraph.GetText?.() || '').join('\n')}\n${drawingText}`,
      series: series.map((item: any) => {
        const colors = item.spPr?.Fill?.fill?.colors || [];
        return colors.map((stop: any) => stop.color?.color?.RGBA || stop.color?.RGBA).filter(Boolean);
      }),
    };
  });
  expect(model.text).toContain('Welcome to ONLYOFFICE Online Editors');
  expect(model.text).toContain('Click here to see the video comparison');
  expect(model.series).toHaveLength(5);
  expect(failures).toEqual([]);
});

test('RTF source retains the complete readable document body instead of only reaching READY', async ({ page }) => {
  const failures = collectPageFailures(page);
  const params = new URLSearchParams({
    scenario: 'local-file',
    type: 'rtf',
    fixtureUrl: '/fixtures/regressions/example-document-title.rtf',
    fixtureName: 'Example Document Title.rtf',
  });
  await page.goto(`/save-e2e.html?${params}`);
  await page.waitForFunction(
    () => {
      const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
      return status?.ready === true || Boolean(status?.error);
    },
    null,
    { timeout: 25_000 },
  );
  expect((await getStatus(page)).error).toBe('');
  const editorFrame = await getWordEditorFrame(page);
  const body = await editorFrame.evaluate(() => {
    const scope = globalThis as typeof globalThis & { editor: any };
    const document = scope.editor.WordControl.m_oLogicDocument;
    const paragraphs = document.GetAllParagraphs?.() || [];
    const drawingText = (document.DrawingObjects?.drawingObjects || [])
      .flatMap((drawing: any) => {
        const graphic = drawing.GraphicObj;
        const shapeParagraphs =
          graphic?.txBody?.content?.GetAllParagraphs?.() || graphic?.textBoxContent?.GetAllParagraphs?.() || [];
        return shapeParagraphs.map((paragraph: any) => paragraph.GetText?.() || '');
      })
      .join('\n');
    return {
      paragraphCount: paragraphs.length,
      text: `${paragraphs.map((paragraph: any) => paragraph.GetText?.() || '').join('\n')}\n${drawingText}`,
    };
  });
  expect(body.paragraphCount).toBeGreaterThan(10);
  expect(body.text).toContain('Welcome to ONLYOFFICE Online Editors');
  expect(body.text).toContain('Click here to see the video comparison');
  expect(failures).toEqual([]);
});

async function getWordEditorFrame(page: Page): Promise<Frame> {
  await expect
    .poll(() => page.frames().find((frame) => frame.url().includes('/documenteditor/main/index.html'))?.url())
    .toContain('/documenteditor/main/index.html');
  const frame = page.frames().find((candidate) => candidate.url().includes('/documenteditor/main/index.html'));
  if (!frame) throw new Error('Word editor frame was not created');
  return frame;
}

function collectPageFailures(page: Page): string[] {
  const failures: string[] = [];
  page.on('pageerror', (error) => failures.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') failures.push(message.text());
  });
  page.on('dialog', async (dialog) => {
    failures.push(`Unexpected dialog: ${dialog.message()}`);
    await dialog.dismiss().catch(() => undefined);
  });
  return failures;
}

async function getStatus(page: Page): Promise<SaveE2EStatus> {
  return page.evaluate(() => {
    const api = (
      window as Window & {
        __ONLYOFFICE_SAVE_E2E__?: { getStatus: () => SaveE2EStatus };
      }
    ).__ONLYOFFICE_SAVE_E2E__;
    if (!api) throw new Error('Save E2E controller is not installed');
    return api.getStatus();
  });
}
