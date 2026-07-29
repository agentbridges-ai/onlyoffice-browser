export interface FontMetadataManifest {
  allFonts?: string;
  defaultFont?: string;
  defaultFonts?: string[];
  builtInFonts?: string[];
  fontFamilies?: Array<{ name: string; paths: string[] }>;
}

export interface FontMetadataFallbackConfig {
  fallbackFamilyName: string;
  unavailableFamilyNames: string[];
  visibleFamilyNames: string[];
}

export function resolveFontMetadataFallbackConfig(
  manifest: FontMetadataManifest,
  downloadedFontPaths: Iterable<string>,
  fallbackFamilyName = 'DengXian',
): FontMetadataFallbackConfig {
  const availablePaths = new Set([
    ...(manifest.defaultFonts || (manifest.defaultFont ? [manifest.defaultFont] : [])),
    ...(manifest.builtInFonts || []),
    ...downloadedFontPaths,
  ]);
  const visibleFamilyNames: string[] = [];
  const unavailableFamilyNames: string[] = [];

  for (const family of manifest.fontFamilies || []) {
    const isAvailable = family.paths.length > 0 && family.paths.every((path) => availablePaths.has(path));
    (isAvailable ? visibleFamilyNames : unavailableFamilyNames).push(family.name);
  }

  return {
    fallbackFamilyName,
    unavailableFamilyNames,
    visibleFamilyNames,
  };
}

/**
 * AllFonts.js is the authoritative input used to construct OnlyOffice's font
 * indexes. Remap unavailable families there, before sdk-all.js creates glyph
 * and zoom caches. A later GetFontFileWeb wrapper is too late: a canvas can
 * retain glyph ids calculated for the original family and then load fallback
 * bytes after a zoom, which renders different characters.
 */
export function buildAllFontsMetadataFallbackBootstrap(config: FontMetadataFallbackConfig): string {
  const serializedConfig = JSON.stringify(config).replaceAll('<', '\\u003c');
  return `\n;(() => {
  const config = ${serializedConfig};
  const infos = window["__fonts_infos"];
  if (!Array.isArray(infos)) return;
  const fallback = infos.find((info) => Array.isArray(info) && info[0] === config.fallbackFamilyName);
  if (!fallback) return;
  const unavailable = new Set(config.unavailableFamilyNames);
  for (let index = 0; index < infos.length; index += 1) {
    const info = infos[index];
    if (Array.isArray(info) && unavailable.has(info[0])) {
      infos[index] = [info[0], ...fallback.slice(1)];
    }
  }
  window["__fonts_visible_names"] = config.visibleFamilyNames.slice();
  window["__onlyOfficeBrowserFontMetadataFallback"] = config;
})();\n`;
}
