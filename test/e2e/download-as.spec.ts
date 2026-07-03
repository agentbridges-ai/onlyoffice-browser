import { expect, test } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';

test.setTimeout(4 * 60_000);

type SourceType = 'doc' | 'docx' | 'xls' | 'xlsx' | 'ppt' | 'pptx';
type DownloadKind =
  | 'csv'
  | 'fb2'
  | 'html'
  | 'image-jpeg'
  | 'image-png'
  | 'image-zip'
  | 'markdown'
  | 'pdf'
  | 'rtf'
  | 'text'
  | 'zip';
type DownloadTarget = {
  ext: string;
  outputExt?: string;
  kind: DownloadKind;
};
type OfficeFixtureCase = {
  filePath: string;
  fileName: string;
  sourceType: SourceType;
  baseName: string;
  multiPage: boolean;
};

const HOST_URL = encodeURIComponent('http://localhost:4173/office-host.html');
const OFFICE_FIXTURE_ROOT = '/Users/xy/Documents/office';
const SUPPORTED_SOURCE_EXTENSIONS = new Set<SourceType>(['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx']);
const DOWNLOAD_AS_TIMEOUT_MS = 15_000;
const DOWNLOAD_AS_DIALOG_POLL_MS = 250;

const EDITOR_FRAME_PART_BY_SOURCE: Record<SourceType, string> = {
  doc: '/documenteditor/',
  docx: '/documenteditor/',
  ppt: '/presentationeditor/',
  pptx: '/presentationeditor/',
  xls: '/spreadsheeteditor/',
  xlsx: '/spreadsheeteditor/',
};

const DOWNLOAD_AS_TARGETS_BY_SOURCE: Record<SourceType, DownloadTarget[]> = {
  doc: wordTargets(),
  docx: wordTargets(),
  ppt: presentationTargets(),
  pptx: presentationTargets(),
  xls: spreadsheetTargets(),
  xlsx: spreadsheetTargets(),
};

function collectOfficeFixtureCases(root: string): OfficeFixtureCase[] {
  const files: string[] = [];
  const visit = (dir: string) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const entryPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        visit(entryPath);
        continue;
      }
      if (!entry.isFile()) continue;
      if (entry.name === '.DS_Store' || entry.name.startsWith('~$')) continue;
      files.push(entryPath);
    }
  };

  visit(root);
  return files
    .map((filePath) => {
      const fileName = path.basename(filePath);
      const sourceType = path.extname(fileName).slice(1).toLowerCase() as SourceType;
      if (!SUPPORTED_SOURCE_EXTENSIONS.has(sourceType)) return null;
      const baseName = fileName.replace(/\.[^/.]+$/, '');
      return {
        filePath,
        fileName,
        sourceType,
        baseName,
        multiPage: /multi_page/i.test(baseName),
      } satisfies OfficeFixtureCase;
    })
    .filter((value): value is OfficeFixtureCase => Boolean(value))
    .sort((left, right) => left.filePath.localeCompare(right.filePath));
}

const OFFICE_FIXTURES = fs.existsSync(OFFICE_FIXTURE_ROOT) ? collectOfficeFixtureCases(OFFICE_FIXTURE_ROOT) : [];

function wordTargets(): DownloadTarget[] {
  return [
    { ext: 'docx', kind: 'zip' },
    { ext: 'pdf', kind: 'pdf' },
    { ext: 'odt', kind: 'zip' },
    { ext: 'dotx', kind: 'zip' },
    { ext: 'pdfa', outputExt: 'pdf', kind: 'pdf' },
    { ext: 'ott', kind: 'zip' },
    { ext: 'md', kind: 'markdown' },
    { ext: 'rtf', kind: 'rtf' },
    { ext: 'txt', kind: 'text' },
    { ext: 'fb2', kind: 'fb2' },
    { ext: 'epub', kind: 'zip' },
    { ext: 'html', kind: 'html' },
    { ext: 'jpg', kind: 'image-jpeg' },
    { ext: 'png', kind: 'image-png' },
  ];
}

function spreadsheetTargets(): DownloadTarget[] {
  return [
    { ext: 'xlsx', kind: 'zip' },
    { ext: 'ods', kind: 'zip' },
    { ext: 'csv', kind: 'csv' },
    { ext: 'pdf', kind: 'pdf' },
    { ext: 'xltx', kind: 'zip' },
    { ext: 'ots', kind: 'zip' },
    { ext: 'xlsb', kind: 'zip' },
    { ext: 'pdfa', outputExt: 'pdf', kind: 'pdf' },
    { ext: 'jpg', kind: 'image-jpeg' },
    { ext: 'png', kind: 'image-png' },
  ];
}

function presentationTargets(): DownloadTarget[] {
  return [
    { ext: 'pptx', kind: 'zip' },
    { ext: 'ppsx', kind: 'zip' },
    { ext: 'pdf', kind: 'pdf' },
    { ext: 'odp', kind: 'zip' },
    { ext: 'potx', kind: 'zip' },
    { ext: 'pdfa', outputExt: 'pdf', kind: 'pdf' },
    { ext: 'otp', kind: 'zip' },
    { ext: 'jpg', kind: 'image-jpeg' },
    { ext: 'png', kind: 'image-png' },
  ];
}

function targetsForFixture(fixture: OfficeFixtureCase): DownloadTarget[] {
  return DOWNLOAD_AS_TARGETS_BY_SOURCE[fixture.sourceType].map((target) => {
    if ((target.ext === 'jpg' || target.ext === 'png') && fixture.multiPage) {
      return { ...target, outputExt: 'zip', kind: 'image-zip' };
    }
    return target;
  });
}

function expectedDownloadName(fixture: OfficeFixtureCase, target: DownloadTarget): string {
  if (target.kind === 'image-zip') {
    const outputExt = target.ext === 'jpeg' ? 'jpg' : target.ext;
    return `${fixture.baseName}_${fixture.sourceType}_${outputExt}.zip`;
  }
  return `${fixture.baseName}.${target.outputExt || target.ext}`;
}

async function waitForSaveE2EReady(page: import('@playwright/test').Page): Promise<void> {
  await page.waitForFunction(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus().ready === true, null, {
    timeout: 90_000,
  });
  const status = await page.evaluate(() => window.__ONLYOFFICE_SAVE_E2E__?.getStatus());
  expect(status?.error).toBe('');
}

async function readDownload(download: import('@playwright/test').Download, expectedName: string): Promise<Buffer> {
  expect(download.suggestedFilename()).toBe(expectedName);
  expect(await download.failure()).toBeNull();

  const stream = await download.createReadStream();
  expect(stream).toBeTruthy();

  const chunks: Buffer[] = [];
  for await (const chunk of stream!) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }

  return Buffer.concat(chunks);
}

async function readDownloadWithExpectedNames(
  download: import('@playwright/test').Download,
  expectedNames: string[],
): Promise<{ content: Buffer; fileName: string }> {
  const fileName = download.suggestedFilename();
  expect(expectedNames).toContain(fileName);
  expect(await download.failure()).toBeNull();

  const stream = await download.createReadStream();
  expect(stream).toBeTruthy();

  const chunks: Buffer[] = [];
  for await (const chunk of stream!) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }

  return { content: Buffer.concat(chunks), fileName };
}

function expectNonEmpty(content: Buffer): void {
  expect(content.byteLength).toBeGreaterThan(0);
}

async function expectPdfDownload(download: import('@playwright/test').Download, expectedName: string): Promise<void> {
  const content = await readDownload(download, expectedName);
  const header = content.subarray(0, 8);
  const latin1 = content.toString('latin1');
  expectNonEmpty(content);
  expect(Array.from(header.slice(0, 4))).toEqual([0x25, 0x50, 0x44, 0x46]);
  expect(latin1).not.toMatch(/\/Type\s*\/Pages\b[\s\S]{0,300}\/Count\s+0\b/);
  expect(latin1).not.toMatch(/\/Count\s+0\b[\s\S]{0,300}\/Type\s*\/Pages\b/);
}

async function expectZipDownload(download: import('@playwright/test').Download, expectedName: string): Promise<void> {
  const content = await readDownload(download, expectedName);
  expectNonEmpty(content);
  expect(Array.from(content.subarray(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
}

function getZipEntryNames(content: Buffer): string[] {
  const names: string[] = [];
  let offset = 0;
  while (offset <= content.byteLength - 46) {
    if (content.readUInt32LE(offset) !== 0x02014b50) {
      offset += 1;
      continue;
    }

    const fileNameLength = content.readUInt16LE(offset + 28);
    const extraLength = content.readUInt16LE(offset + 30);
    const commentLength = content.readUInt16LE(offset + 32);
    const nameStart = offset + 46;
    const nameEnd = nameStart + fileNameLength;
    if (nameEnd > content.byteLength) break;
    names.push(content.subarray(nameStart, nameEnd).toString('utf8'));
    offset = nameEnd + extraLength + commentLength;
  }
  return names;
}

async function expectImageZipDownload(
  download: import('@playwright/test').Download,
  target: DownloadTarget,
  expectedName: string,
): Promise<void> {
  const content = await readDownload(download, expectedName);
  expectNonEmpty(content);
  expect(Array.from(content.subarray(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
  const names = getZipEntryNames(content);
  const imagePattern = target.ext === 'png' ? /\.png$/i : /\.(jpe?g)$/i;
  expect(names.some((name) => imagePattern.test(name)), `image entries in ${expectedName}: ${names.join(', ')}`).toBe(true);
}

async function expectMarkdownDownload(
  download: import('@playwright/test').Download,
  expectedName: string,
  fixture: OfficeFixtureCase,
): Promise<void> {
  const baseName = expectedName.replace(/\.md$/i, '');
  const { content, fileName } = await readDownloadWithExpectedNames(download, [
    `${baseName}.md`,
    `${fixture.baseName}_${fixture.sourceType}_md.zip`,
  ]);
  expectNonEmpty(content);

  if (/\.zip$/i.test(fileName)) {
    expect(Array.from(content.subarray(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
    const names = getZipEntryNames(content);
    expect(names.some((name) => /\.md$/i.test(name)), `markdown entries in ${fileName}: ${names.join(', ')}`).toBe(
      true,
    );
    expect(
      names.some((name) => /^assets\/.+\.(png|jpe?g|gif|webp|svg)$/i.test(name)),
      `asset entries in ${fileName}: ${names.join(', ')}`,
    ).toBe(true);
    return;
  }

  const text = content.toString('utf8').replace(/^\ufeff/, '').trimStart();
  expect(text.length).toBeGreaterThan(0);
  expect(text).not.toMatch(/^(DOCY;|XLSY;|PPTY;|RE9D|UEsDB|PK\x03\x04)/);
}

async function expectTextDownload(
  download: import('@playwright/test').Download,
  expectedName: string,
  matcher?: RegExp,
): Promise<void> {
  const content = await readDownload(download, expectedName);
  expectNonEmpty(content);
  expect(content.includes(0)).toBe(false);
  const text = content.toString('utf8').replace(/^\ufeff/, '').trimStart();
  expect(text.length).toBeGreaterThan(0);
  expect(text).not.toMatch(/^(DOCY;|XLSY;|PPTY;|RE9D|UEsDB|PK\x03\x04)/);
  if (matcher) expect(text).toMatch(matcher);
}

async function expectRtfDownload(download: import('@playwright/test').Download, expectedName: string): Promise<void> {
  const content = await readDownload(download, expectedName);
  expectNonEmpty(content);
  const latin1 = content.toString('latin1').trimStart();
  const utf16le = content.toString('utf16le').replace(/^\ufeff/, '').trimStart();
  expect(latin1.startsWith('{\\rtf') || utf16le.startsWith('{\\rtf')).toBe(true);
}

async function expectJpegDownload(download: import('@playwright/test').Download, expectedName: string): Promise<void> {
  const content = await readDownload(download, expectedName);
  expectNonEmpty(content);
  expect(Array.from(content.subarray(0, 2))).toEqual([0xff, 0xd8]);
  expect(Array.from(content.subarray(-2))).toEqual([0xff, 0xd9]);
}

async function expectPngDownload(download: import('@playwright/test').Download, expectedName: string): Promise<void> {
  const content = await readDownload(download, expectedName);
  expectNonEmpty(content);
  expect(Array.from(content.subarray(0, 8))).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
}

async function expectDownloadByKind(
  download: import('@playwright/test').Download,
  target: DownloadTarget,
  expectedName: string,
  fixture: OfficeFixtureCase,
): Promise<void> {
  if (target.kind === 'pdf') {
    await expectPdfDownload(download, expectedName);
    return;
  }
  if (target.kind === 'zip') {
    await expectZipDownload(download, expectedName);
    return;
  }
  if (target.kind === 'rtf') {
    await expectRtfDownload(download, expectedName);
    return;
  }
  if (target.kind === 'html') {
    await expectTextDownload(download, expectedName, /<!doctype\s+html|<html\b|<body\b|<[a-z][\w:-]*\b/i);
    return;
  }
  if (target.kind === 'fb2') {
    await expectTextDownload(download, expectedName, /fictionbook/i);
    return;
  }
  if (target.kind === 'image-zip') {
    await expectImageZipDownload(download, target, expectedName);
    return;
  }
  if (target.kind === 'markdown') {
    await expectMarkdownDownload(download, expectedName, fixture);
    return;
  }
  if (target.kind === 'image-jpeg') {
    await expectJpegDownload(download, expectedName);
    return;
  }
  if (target.kind === 'image-png') {
    await expectPngDownload(download, expectedName);
    return;
  }
  await expectTextDownload(download, expectedName);
}

async function clickIfVisible(locator: import('@playwright/test').Locator): Promise<boolean> {
  try {
    if (!(await locator.isVisible({ timeout: 250 }))) return false;
    await locator.click({ timeout: 5_000, force: true });
    return true;
  } catch {
    return false;
  }
}

async function clickDownloadDialogAction(frame: import('@playwright/test').Frame): Promise<boolean> {
  const labels = [
    'Save & Download',
    'Save and Download',
    '保存并下载',
    '保存 & 下载',
    'Download',
    '下载',
    'OK',
    'Ok',
    '确定',
    'Apply',
    '应用',
  ];
  const clicked = await frame.evaluate((candidateLabels) => {
    function isVisible(element: Element): boolean {
      const style = window.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        Number(style.opacity || '1') !== 0 &&
        rect.width > 0 &&
        rect.height > 0
      );
    }

    const normalizedLabels = candidateLabels.map((label) => label.toLowerCase());
    const candidates = Array.from(
      document.querySelectorAll<HTMLElement>('button, [role="button"], .btn, .button, a'),
    )
      .filter((element) => {
        const text = (element.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
        return normalizedLabels.includes(text) && isVisible(element);
      })
      .sort((left, right) => {
        const leftRect = left.getBoundingClientRect();
        const rightRect = right.getBoundingClientRect();
        return leftRect.width * leftRect.height - rightRect.width * rightRect.height;
      });

    const target = candidates[0];
    if (!target) return false;
    target.click();
    return true;
  }, labels);
  if (clicked) return true;

  for (const pattern of [/Save & Download|保存并下载|保存 & 下载/i, /^Download$|^下载$/i, /^OK$|^Ok$|^确定$/i, /^Apply$|^应用$/i]) {
    if (await clickIfVisible(frame.getByRole('button', { name: pattern }).first())) return true;
    if (await clickIfVisible(frame.getByText(pattern).first())) return true;
  }
  return false;
}

async function clickVisibleText(frame: import('@playwright/test').Frame, labels: string[]): Promise<boolean> {
  const clicked = await frame.evaluate((candidateLabels) => {
    function isVisible(element: Element): boolean {
      const style = window.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        Number(style.opacity || '1') !== 0 &&
        rect.width > 0 &&
        rect.height > 0
      );
    }

    const normalizedLabels = candidateLabels.map((label) => label.toLowerCase());
    const candidates = Array.from(
      document.querySelectorAll<HTMLElement>('button, a, [role="button"], [role="tab"], li, span, div'),
    )
      .filter((element) => {
        const text = (element.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
        return normalizedLabels.includes(text) && isVisible(element);
      })
      .sort((left, right) => {
        const leftRect = left.getBoundingClientRect();
        const rightRect = right.getBoundingClientRect();
        return leftRect.width * leftRect.height - rightRect.width * rightRect.height;
      });

    const target = candidates[0];
    if (!target) return false;
    target.click();
    return true;
  }, labels);

  return clicked;
}

async function isDownloadFormatVisible(frame: import('@playwright/test').Frame, ext: string): Promise<boolean> {
  return frame.evaluate((targetExt) => {
    function isVisible(element: Element): boolean {
      const style = window.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        Number(style.opacity || '1') !== 0 &&
        rect.width > 0 &&
        rect.height > 0
      );
    }

    return Array.from(document.querySelectorAll<HTMLElement>('.btn-doc-format[format]')).some(
      (button) => Boolean(button.querySelector(`.svg-format-${targetExt}`)) && isVisible(button),
    );
  }, ext);
}

async function waitForDownloadFormat(frame: import('@playwright/test').Frame, ext: string): Promise<void> {
  await frame.waitForFunction(
    (targetExt) => {
      function isVisible(element: Element): boolean {
        const style = window.getComputedStyle(element);
        const rect = element.getBoundingClientRect();
        return (
          style.display !== 'none' &&
          style.visibility !== 'hidden' &&
          Number(style.opacity || '1') !== 0 &&
          rect.width > 0 &&
          rect.height > 0
        );
      }

	      return Array.from(document.querySelectorAll<HTMLElement>('.btn-doc-format[format]')).some(
	        (button) => Boolean(button.querySelector(`.svg-format-${targetExt}`)) && isVisible(button),
	      );
	    },
	    ext,
	    { timeout: 5_000 },
	  );
}

async function clickDownloadFormat(frame: import('@playwright/test').Frame, ext: string): Promise<void> {
  const clicked = await frame.evaluate((targetExt) => {
    function isVisible(element: Element): boolean {
      const style = window.getComputedStyle(element);
      const rect = element.getBoundingClientRect();
      return (
        style.display !== 'none' &&
        style.visibility !== 'hidden' &&
        Number(style.opacity || '1') !== 0 &&
        rect.width > 0 &&
        rect.height > 0
      );
    }

    const button = Array.from(document.querySelectorAll<HTMLElement>('.btn-doc-format[format]')).find(
      (candidate) => Boolean(candidate.querySelector(`.svg-format-${targetExt}`)) && isVisible(candidate),
    );
    if (!button) return false;
    button.click();
    return true;
  }, ext);
  expect(clicked, `visible Download As ${ext.toUpperCase()} tile`).toBe(true);
}

async function openDownloadAsPanel(
  page: import('@playwright/test').Page,
  frame: import('@playwright/test').Frame,
  target: DownloadTarget,
): Promise<void> {
  const deadline = Date.now() + 10_000;

  while (Date.now() < deadline) {
    if (await isDownloadFormatVisible(frame, target.ext)) return;

    await clickVisibleText(frame, ['File', '文件']);
    await page.waitForTimeout(150);
    if (await isDownloadFormatVisible(frame, target.ext)) return;

    await clickVisibleText(frame, ['Download As', 'Download as', '下载为']);
    await page.waitForTimeout(150);
    if (await isDownloadFormatVisible(frame, target.ext)) return;
  }

  throw new Error(`Download As panel did not show ${target.ext.toUpperCase()} within 10000ms`);
}

async function triggerDownloadAs(
  page: import('@playwright/test').Page,
  frame: import('@playwright/test').Frame,
  target: DownloadTarget,
): Promise<import('@playwright/test').Download> {
  await openDownloadAsPanel(page, frame, target);
  await waitForDownloadFormat(frame, target.ext);

  let downloadTimeoutMessage = '';
  const downloadPromise = page
    .waitForEvent('download', { timeout: DOWNLOAD_AS_TIMEOUT_MS })
    .catch((error: Error) => {
      downloadTimeoutMessage = error.message;
      return null;
    });
  await clickDownloadFormat(frame, target.ext);

  let conversionError: Error | null = null;
  const conversionErrorPromise = page
    .waitForFunction(
      () => {
        const status = window.__ONLYOFFICE_SAVE_E2E__?.getStatus();
        return status?.error ? status.error : null;
      },
      null,
      { timeout: DOWNLOAD_AS_TIMEOUT_MS },
    )
    .then(async (handle) => {
      const message = await handle.jsonValue();
      conversionError = new Error(String(message || 'Download As conversion failed'));
    })
    .catch(() => undefined);

  const deadline = Date.now() + DOWNLOAD_AS_TIMEOUT_MS;
  while (Date.now() < deadline) {
    const result = await Promise.race([
      downloadPromise.then((download) => (download ? { download } : null)),
      page.waitForTimeout(DOWNLOAD_AS_DIALOG_POLL_MS).then(() => null),
    ]);
    if (result?.download) return result.download;
    if (conversionError) throw conversionError;
    await clickDownloadDialogAction(frame);
  }

  const finalDownload = await downloadPromise;
  await conversionErrorPromise;
  if (finalDownload) return finalDownload;
  if (conversionError) throw conversionError;
  throw new Error(
    `Download As ${target.ext.toUpperCase()} did not produce a download within ${DOWNLOAD_AS_TIMEOUT_MS}ms${
      downloadTimeoutMessage ? `: ${downloadTimeoutMessage}` : ''
    }`,
  );
}

test.skip(OFFICE_FIXTURES.length === 0, `No Office fixtures found in ${OFFICE_FIXTURE_ROOT}`);

for (const fixture of OFFICE_FIXTURES) {
  const relativeFixturePath = path.relative(OFFICE_FIXTURE_ROOT, fixture.filePath);
  for (const target of targetsForFixture(fixture)) {
    test(`${relativeFixturePath} -> ${target.ext} built-in Download As`, async ({ page }) => {
      await page.route('**/__office-fixture__/**', async (route) => {
        await route.fulfill({
          body: await fs.promises.readFile(fixture.filePath),
          contentType: 'application/octet-stream',
        });
      });

      const fixtureUrl = `/__office-fixture__/${encodeURIComponent(fixture.fileName)}`;
      await page.goto(
        `/save-e2e.html?scenario=local-file&type=${fixture.sourceType}&hostUrl=${HOST_URL}&fixtureUrl=${encodeURIComponent(
          fixtureUrl,
        )}&fixtureName=${encodeURIComponent(fixture.fileName)}`,
      );
      await waitForSaveE2EReady(page);

      const editorFrame = page
        .frames()
        .find((frame) => frame.url().includes(EDITOR_FRAME_PART_BY_SOURCE[fixture.sourceType]));
      expect(editorFrame, `${fixture.fileName} editor frame`).toBeTruthy();

      const download = await triggerDownloadAs(page, editorFrame!, target);
      await expectDownloadByKind(download, target, expectedDownloadName(fixture, target), fixture);
    });
  }
}
