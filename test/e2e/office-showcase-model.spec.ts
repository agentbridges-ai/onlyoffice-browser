import { expect, type Frame, type Page, test } from '@playwright/test';

type SaveE2EStatus = {
  ready: boolean;
  error: string;
  state: { status: string } | null;
};

test('Word showcase exposes its authored content and drawing objects in the editor model', async ({ page }) => {
  const failures = collectPageFailures(page);
  await openFixture(page, 'docx', '/fixtures/regressions/office-preview-showcase.docx', 'ONLYOFFICE Word Preview Showcase.docx');
  const frame = await getEditorFrame(page, '/documenteditor/main/index.html');
  const model = await frame.evaluate(() => {
    const scope = globalThis as typeof globalThis & { editor?: any; AscCommon?: any };
    const document = scope.editor?.WordControl?.m_oLogicDocument;
    const paragraphs = document?.GetAllParagraphs?.() || [];
    const drawings = document?.DrawingObjects?.drawingObjects || [];
    const graphics = drawings.map((drawing: any) => drawing.GraphicObj).filter(Boolean);
    const chart = graphics.find((graphic: any) => graphic.isChart?.());
    const imageResults = graphics
      .map((graphic: any) => graphic.getImageUrl?.())
      .filter(Boolean)
      .map((imageName: string) => {
        const resourceUrl = scope.AscCommon?.g_oDocumentUrls?.urls?.[`media/${imageName}`];
        const image = resourceUrl
          ? (scope.editor?.ImageLoader?.map_image_index?.[resourceUrl] ||
              scope.AscCommon?.g_image_loader?.map_image_index?.[resourceUrl])?.Image
          : undefined;
        return {
          imageName,
          complete: Boolean(image?.complete),
          width: image?.naturalWidth || 0,
          height: image?.naturalHeight || 0,
        };
      });
    return {
      text: paragraphs.map((paragraph: any) => paragraph.GetText?.() || '').join('\n'),
      paragraphCount: paragraphs.length,
      drawingTypes: graphics.map((graphic: any) => graphic.constructor?.name || ''),
      chartSeries: chart?.chart?.plotArea?.charts?.[0]?.series?.length || 0,
      decodedImages: imageResults,
    };
  });

  for (const marker of ['W01', 'W02', 'W03', 'W04', 'W05', 'W06', 'W07', 'W08', 'W09', 'W10', 'W11', 'W12', 'W13', 'W14', 'W15', 'W16', 'W17/W19', 'W18', 'W20', 'W21']) {
    expect(model.text, `missing Word model marker ${marker}`).toContain(marker);
  }
  expect(model.paragraphCount).toBeGreaterThan(40);
  expect(model.drawingTypes.length).toBeGreaterThanOrEqual(5);
  expect(model.chartSeries).toBe(5);
  expect(model.decodedImages.length).toBeGreaterThanOrEqual(2);
  expect(
    model.decodedImages.every(
      (image: { complete: boolean; width: number; height: number }) =>
        image.complete && image.width > 0 && image.height > 0,
    ),
  ).toBe(true);
  expect(failures).toEqual([]);
});

test('PowerPoint showcase exposes every slide, marker and native object class in the editor model', async ({ page }) => {
  const failures = collectPageFailures(page);
  await openFixture(page, 'pptx', '/fixtures/regressions/office-preview-showcase.pptx', 'ONLYOFFICE Presentation Preview Showcase.pptx');
  const frame = await getEditorFrame(page, '/presentationeditor/main/index.html');
  const model = await frame.evaluate(() => {
    const scope = globalThis as typeof globalThis & { editor?: any };
    const presentation = scope.editor?.WordControl?.m_oLogicDocument;
    const slides = presentation?.Slides || [];
    const summary = { text: '', shapeCount: 0, chartCount: 0, groupCount: 0, tableCount: 0, imageCount: 0, transitionCount: 0, timingCount: 0 };
    const visit = (shape: any) => {
      summary.shapeCount += 1;
      const type = shape?.constructor?.name || '';
      if (type === 'CGroupShape' || Array.isArray(shape?.spTree)) summary.groupCount += 1;
      if (shape?.isChart?.() || shape?.graphicObject?.isChart?.()) summary.chartCount += 1;
      if (type === 'CGraphicFrame' && shape?.graphicObject?.constructor?.name === 'CTable') summary.tableCount += 1;
      if (shape?.getImageUrl?.()) summary.imageCount += 1;
      const paragraphs = shape?.txBody?.content?.GetAllParagraphs?.() || shape?.textBoxContent?.GetAllParagraphs?.() || [];
      summary.text += `\n${paragraphs.map((paragraph: any) => paragraph.GetText?.() || '').join('\n')}`;
      for (const child of shape?.spTree || []) visit(child);
    };
    for (const slide of slides) {
      if (slide?.transition || slide?.Transition) summary.transitionCount += 1;
      if (slide?.timing || slide?.Timing) summary.timingCount += 1;
      for (const shape of slide?.cSld?.spTree || []) visit(shape);
    }
    return { slideCount: slides.length, ...summary };
  });

  expect(model.slideCount).toBe(9);
  for (const marker of ['P01', 'P02/P03', 'P04/P05', 'P06', 'P07', 'P08', 'P09/P10', 'P11/P12', 'P13']) {
    expect(model.text, `missing PowerPoint model marker ${marker}`).toContain(marker);
  }
  expect(model.shapeCount).toBeGreaterThan(40);
  expect(model.chartCount).toBe(3);
  expect(model.groupCount).toBeGreaterThanOrEqual(2);
  expect(model.tableCount).toBeGreaterThanOrEqual(1);
  expect(model.imageCount).toBeGreaterThanOrEqual(2);
  expect(model.transitionCount).toBe(9);
  expect(model.timingCount).toBeGreaterThanOrEqual(1);
  expect(failures).toEqual([]);
});

test('Spreadsheet showcase exposes sheets, formulas, charts and native-object sheet in the workbook model', async ({ page }) => {
  const failures = collectPageFailures(page);
  await openFixture(page, 'xlsx', '/fixtures/regressions/office-preview-showcase.xlsx', 'ONLYOFFICE Spreadsheet Preview Showcase.xlsx');
  const frame = await getEditorFrame(page, '/spreadsheeteditor/main/index.html');
  const model = await frame.evaluate(() => {
    const scope = globalThis as typeof globalThis & { Asc?: { editor?: any }; editor?: any };
    const api = scope.Asc?.editor || scope.editor;
    const workbook = api?.wbModel;
    const sheets = workbook?.aWorksheets || [];
    const readCell = (sheet: any, row: number, col: number) => {
      const cell = sheet?.getCell3?.(row, col);
      return { value: cell?.getValue?.() ?? '', formula: cell?.getFormula?.() ?? '' };
    };
    return {
      sheetNames: sheets.map((sheet: any) => sheet.getName?.() || sheet.sName || ''),
      hiddenSheets: sheets.filter((sheet: any) => sheet.getHidden?.() || sheet.hidden).map((sheet: any) => sheet.getName?.() || sheet.sName || ''),
      drawingCounts: sheets.map((sheet: any) => sheet.Drawings?.length || 0),
      chartCount: sheets.reduce(
        (count: number, sheet: any) => count + (sheet.Drawings || []).filter((drawing: any) => drawing.graphicObject?.isChart?.()).length,
        0,
      ),
      overviewTitle: readCell(sheets.find((sheet: any) => (sheet.getName?.() || sheet.sName) === 'Overview'), 1, 1),
      overviewFormula: readCell(sheets.find((sheet: any) => (sheet.getName?.() || sheet.sName) === 'Overview'), 6, 1),
      nativeTitle: readCell(sheets.find((sheet: any) => (sheet.getName?.() || sheet.sName) === 'Native PivotTable and Slicer'), 0, 0),
    };
  });

  expect(model.sheetNames).toEqual(['Overview', 'Data', 'Formats', 'Charts', 'Sparklines', 'RTL & CJK', 'Reference', 'Native PivotTable and Slicer']);
  expect(model.hiddenSheets).toContain('Reference');
  expect(model.overviewTitle.value).toContain('Spreadsheet Preview Showcase');
  expect(model.overviewFormula.formula).toContain('SUM');
  expect(model.nativeTitle.value).toBe('Status');
  expect(model.chartCount).toBe(6);
  expect(model.drawingCounts.reduce((sum: number, count: number) => sum + count, 0)).toBeGreaterThanOrEqual(9);
  expect(failures).toEqual([]);
});

async function openFixture(page: Page, type: string, fixtureUrl: string, fixtureName: string): Promise<void> {
  const params = new URLSearchParams({ scenario: 'local-file', type, fixtureUrl, fixtureName });
  await page.goto(`/save-e2e.html?${params}`);
  await page.waitForFunction(
    () => {
      const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
      return status?.ready === true || Boolean(status?.error);
    },
    null,
    { timeout: 25_000 },
  );
  const status = await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus() as SaveE2EStatus);
  expect(status.error).toBe('');
  expect(status.ready).toBe(true);
  expect(status.state?.status).toBe('ready');
}

async function getEditorFrame(page: Page, path: string): Promise<Frame> {
  await expect.poll(() => page.frames().find((frame) => frame.url().includes(path))?.url()).toContain(path);
  const frame = page.frames().find((candidate) => candidate.url().includes(path));
  if (!frame) throw new Error(`Editor frame was not created: ${path}`);
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
