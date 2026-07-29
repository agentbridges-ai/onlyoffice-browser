import { BASE_PATH } from './document-utils';

export const GENERATED_FONT_ASSETS_MANIFEST = 'onlyoffice-browser-font-assets.json';
export const GENERATED_FONT_SOURCE_MAP = 'onlyoffice-browser-font-source-map.json';

export interface GeneratedFontSourceMapEntry {
  source: string;
  file: string;
}

export interface GeneratedFontSourceMap {
  fonts: GeneratedFontSourceMapEntry[];
}

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
  fontSourceMap?: string;
  defaultFont?: string;
  defaultFonts?: string[];
  builtInFonts?: string[];
  fontFamilies?: Array<{ name: string; paths: string[] }>;
}

const FONT_ASSETS_SETUP_HINT =
  'Generate them with `npm run fonts:generate -- --input /path/to/fonts --output .onlyoffice-font-assets`, ' +
  'then serve that directory with `ONLYOFFICE_BROWSER_FONT_ASSETS_DIR=/absolute/path/to/.onlyoffice-font-assets` in dev, ' +
  'or deploy the generated directory at the editor host root in production.';

export function resolveRuntimeAssetCacheMode(hostname: string): RequestCache {
  // Piwork and the Cloudflare wildcard deployment redirect shareable assets to
  // one immutable, build-versioned URL. Distinct editor origins can therefore
  // reuse one browser HTTP-cache entry without revalidation.
  return hostname.endsWith('.office.localhost') ||
    hostname === 'onlyoffice.getpi.work' ||
    /^office-editor-[a-z0-9-]+\.getpi\.work$/.test(hostname)
    ? 'force-cache'
    : 'no-cache';
}

function runtimeAssetCacheMode(): RequestCache {
  return resolveRuntimeAssetCacheMode(globalThis.location.hostname);
}

function normalizeAssetPath(assetPath: string): string {
  return assetPath.replace(/^\/+/, '');
}

export function getRuntimeAssetUrl(assetPath: string): string {
  const baseUrl = new URL(BASE_PATH, globalThis.location.href);
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

function isFontFamilyArray(value: unknown): value is Array<{ name: string; paths: string[] }> {
  return (
    Array.isArray(value) &&
    value.every(
      (family) =>
        family &&
        typeof family === 'object' &&
        typeof (family as { name?: unknown }).name === 'string' &&
        isStringArray((family as { paths?: unknown }).paths),
    )
  );
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
  if (manifest.defaultFonts !== undefined && !isStringArray(manifest.defaultFonts)) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} has invalid defaultFonts`);
  }
  if (manifest.builtInFonts !== undefined && !isStringArray(manifest.builtInFonts)) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} has invalid builtInFonts`);
  }
  if (manifest.fontFamilies !== undefined && !isFontFamilyArray(manifest.fontFamilies)) {
    throw createMissingFontAssetsError(`${GENERATED_FONT_ASSETS_MANIFEST} has invalid fontFamilies`);
  }

  return manifest as GeneratedFontAssetsManifest;
}

/**
 * Resolve picker-visible families from the verified font paths supplied by the
 * embedding application. AllFonts.js remains the glyph-index source, but it is
 * deliberately not the authority for runtime font availability.
 */
export function resolveAvailableFontFamilyNames(
  manifest: GeneratedFontAssetsManifest,
  verifiedFontPaths: string[],
): string[] {
  const availablePaths = new Set([
    ...(manifest.defaultFonts || (manifest.defaultFont ? [manifest.defaultFont] : [])),
    ...(manifest.builtInFonts || []),
    ...verifiedFontPaths,
  ]);
  return (manifest.fontFamilies || [])
    .filter((family) => family.paths.length > 0 && family.paths.every((path) => availablePaths.has(path)))
    .map((family) => family.name);
}

export async function fetchGeneratedFontAssetsManifest(): Promise<GeneratedFontAssetsManifest> {
  const manifestUrl = getRuntimeAssetUrl(GENERATED_FONT_ASSETS_MANIFEST);
  let response: Response;
  try {
    response = await fetch(manifestUrl, { cache: runtimeAssetCacheMode() });
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

export async function fetchGeneratedFontSourceMap(assetPath?: string): Promise<GeneratedFontSourceMap | null> {
  if (!assetPath) return null;

  const response = await fetch(getRuntimeAssetUrl(assetPath), { cache: runtimeAssetCacheMode() }).catch(() => null);
  if (!response || response.status === 404) return null;
  if (!response.ok) {
    throw createMissingFontAssetsError(`${assetPath} returned ${response.status}`);
  }

  let value: Partial<GeneratedFontSourceMap>;
  try {
    value = (await response.json()) as Partial<GeneratedFontSourceMap>;
  } catch {
    return null;
  }

  if (!value || !Array.isArray(value.fonts)) return null;
  return {
    fonts: value.fonts.filter((font): font is GeneratedFontSourceMapEntry =>
      Boolean(
        font &&
        typeof font === 'object' &&
        typeof (font as GeneratedFontSourceMapEntry).source === 'string' &&
        typeof (font as GeneratedFontSourceMapEntry).file === 'string',
      ),
    ),
  };
}

async function assertRuntimeAssetReachable(assetPath: string): Promise<void> {
  const assetUrl = getRuntimeAssetUrl(assetPath);
  let response: Response;
  try {
    response = await fetch(assetUrl, {
      method: 'HEAD',
      // A byte-range probe can poison the browser HTTP cache when the server
      // transparently compresses JavaScript: Chrome may combine the cached
      // identity byte with a later gzip representation. HEAD verifies the
      // immutable object without downloading or caching a partial body.
      cache: 'no-store',
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
  const startupFont =
    manifest.defaultFonts?.[0] ||
    manifest.defaultFont ||
    manifest.builtInFonts?.[0] ||
    manifest.fonts[0];
  await Promise.all([
    assertRuntimeAssetReachable(manifest.allFonts),
    assertRuntimeAssetReachable(manifest.fontSelection),
    ...(manifest.fontSourceMap ? [assertRuntimeAssetReachable(manifest.fontSourceMap)] : []),
    assertRuntimeAssetReachable(manifest.fontThumbnails[0]),
    assertRuntimeAssetReachable(startupFont),
  ]);
  return manifest;
}

export async function fetchRuntimeBinaryAsset(assetPath: string): Promise<Uint8Array> {
  const response = await fetch(getRuntimeAssetUrl(assetPath), { cache: runtimeAssetCacheMode() });
  if (!response.ok) {
    throw createMissingFontAssetsError(`${assetPath} returned ${response.status}`);
  }
  return new Uint8Array(await response.arrayBuffer());
}
