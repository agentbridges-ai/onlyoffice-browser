import { BASE_PATH } from './document-utils';

export const GENERATED_FONT_ASSETS_MANIFEST = 'onlyoffice-browser-font-assets.json';

export interface GeneratedFontAssetsManifest {
  version: number;
  generator?: string;
  image?: string;
  fontSet?: string;
  generatedAt?: string;
  allFonts: string;
  fontSelection: string;
  fontThumbnails: string[];
  fonts: string[];
}

const FONT_ASSETS_SETUP_HINT =
  'Generate them with `npm run fonts:generate -- --input /path/to/fonts --output .onlyoffice-font-assets`, ' +
  'then serve that directory with `ONLYOFFICE_BROWSER_FONT_ASSETS_DIR=/absolute/path/to/.onlyoffice-font-assets` in dev, ' +
  'or deploy the generated directory at the editor host root in production.';

function normalizeAssetPath(assetPath: string): string {
  return assetPath.replace(/^\/+/, '');
}

export function getRuntimeAssetUrl(assetPath: string): string {
  const baseUrl = new URL(BASE_PATH, window.location.href);
  return new URL(normalizeAssetPath(assetPath), baseUrl).href;
}

export function getAssetFileName(assetPath: string): string {
  return normalizeAssetPath(assetPath).split('/').filter(Boolean).pop() || 'font.bin';
}

function createMissingFontAssetsError(detail: string): Error {
  return new Error(`OnlyOffice font assets are required: ${detail}. ${FONT_ASSETS_SETUP_HINT}`);
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === 'string' && item.length > 0);
}

function validateManifest(value: unknown): GeneratedFontAssetsManifest {
  if (!value || typeof value !== 'object') {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} is not a JSON object`);
  }

  const manifest = value as Partial<GeneratedFontAssetsManifest>;
  if (manifest.version !== 1) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} has an unsupported version`);
  }
  if (typeof manifest.allFonts !== 'string' || manifest.allFonts.length === 0) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} is missing allFonts`);
  }
  if (typeof manifest.fontSelection !== 'string' || manifest.fontSelection.length === 0) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} is missing fontSelection`);
  }
  if (!isStringArray(manifest.fontThumbnails) || manifest.fontThumbnails.length === 0) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} is missing fontThumbnails`);
  }
  if (!isStringArray(manifest.fonts) || manifest.fonts.length === 0) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} is missing fonts`);
  }

  return manifest as GeneratedFontAssetsManifest;
}

export async function fetchGeneratedFontAssetsManifest(): Promise<GeneratedFontAssetsManifest> {
  const manifestUrl = getRuntimeAssetUrl(GENERATED_FONT_ASSETS_MANIFEST);
  let response: Response;
  try {
    response = await fetch(manifestUrl, { cache: 'no-cache' });
  } catch (error) {
    throw createMissingFontAssetsError(
      `failed to request ${GENERATED_FONT_ASSETS_MANIFEST}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  if (!response.ok) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} returned ${response.status}`);
  }

  try {
    return validateManifest(await response.json());
  } catch (error) {
    if (error instanceof Error && error.message.startsWith('OnlyOffice font assets are required:')) {
      throw error;
    }
    throw createMissingFontAssetsError(
      `failed to parse ${GENERATED_FONT_ASSETS_MANIFEST}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

async function assertRuntimeAssetReachable(assetPath: string): Promise<void> {
  const assetUrl = getRuntimeAssetUrl(assetPath);
  let response: Response;
  try {
    response = await fetch(assetUrl, {
      cache: 'no-cache',
      headers: {
        Range: 'bytes=0-0',
      },
    });
  } catch (error) {
    throw createMissingFontAssetsError(
      `${assetPath} is not reachable: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  if (response.ok) return;

  throw createMissingFontAssetsError(`${assetPath} returned ${response.status}`);
}

export async function assertGeneratedFontAssetsAvailable(): Promise<GeneratedFontAssetsManifest> {
  const manifest = await fetchGeneratedFontAssetsManifest();
  await Promise.all([
    assertRuntimeAssetReachable(manifest.allFonts),
    assertRuntimeAssetReachable(manifest.fontSelection),
    assertRuntimeAssetReachable(manifest.fontThumbnails[0]),
    assertRuntimeAssetReachable(manifest.fonts[0]),
  ]);
  return manifest;
}

export async function fetchRuntimeBinaryAsset(assetPath: string): Promise<Uint8Array> {
  const response = await fetch(getRuntimeAssetUrl(assetPath), { cache: 'no-cache' });
  if (!response.ok) {
    throw createMissingFontAssetsError(`${assetPath} returned ${response.status}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}
