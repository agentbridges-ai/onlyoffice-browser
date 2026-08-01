#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { COMMON_FONT_PICKER_FAMILIES, REQUIRED_DEFAULT_FONT_PICKER_FAMILIES } from './font-picker-policy.mjs';

const ALL_FONTS_PATH = 'sdkjs/common/AllFonts.js';
const FONT_MANIFEST_PATH = 'onlyoffice-browser-font-assets.json';
const FONT_SOURCE_MAP_PATH = 'onlyoffice-browser-font-source-map.json';

function parseJsArray(source, name) {
  const match = source.match(new RegExp(`window\\["${name}"\\]\\s*=\\s*(\\[[\\s\\S]*?\\]);`));
  if (!match) throw new Error(`Generated AllFonts.js is missing ${name}`);
  return JSON.parse(match[1]);
}

function upsertJsArray(source, name, value) {
  const replacement = `window["${name}"] = ${JSON.stringify(value)};`;
  const pattern = new RegExp(`window\\["${name}"\\]\\s*=\\s*\\[[\\s\\S]*?\\];`);
  if (pattern.test(source)) return source.replace(pattern, replacement);
  const marker = 'window["__fonts_infos"]';
  const index = source.indexOf(marker);
  return index < 0 ? `${replacement}\n${source}` : `${source.slice(0, index)}${replacement}\n${source.slice(index)}`;
}

function revision(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex').slice(0, 16);
}

function updateAssetMetadata(manifest, relativePath, bytes) {
  const asset = Array.isArray(manifest.assets)
    ? manifest.assets.find((candidate) => candidate?.path === relativePath)
    : null;
  if (!asset) {
    throw new Error(`Generated font manifest is missing asset metadata: ${relativePath}`);
  }
  asset.bytes = bytes.byteLength;
  asset.revision = revision(bytes);
}

export function applyFontPickerPolicy(rootDirectory) {
  const root = path.resolve(rootDirectory);
  const allFontsPath = path.join(root, ALL_FONTS_PATH);
  const manifestPath = path.join(root, FONT_MANIFEST_PATH);
  const sourceMapPath = path.join(root, FONT_SOURCE_MAP_PATH);
  const source = fs.readFileSync(allFontsPath, 'utf8');
  const infos = parseJsArray(source, '__fonts_infos');
  const installedNames = new Set(infos.map((info) => info?.[0]).filter((name) => typeof name === 'string'));
  const visibleNames = COMMON_FONT_PICKER_FAMILIES.filter((name) => installedNames.has(name)).sort();
  for (const requiredName of REQUIRED_DEFAULT_FONT_PICKER_FAMILIES) {
    if (!visibleNames.includes(requiredName)) {
      throw new Error(`Generated font package is missing required default family: ${requiredName}`);
    }
  }

  const nextAllFonts = `${upsertJsArray(source, '__fonts_visible_names', visibleNames).trimEnd()}\n`;
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  const familyByName = new Map(
    (Array.isArray(manifest.fontFamilies) ? manifest.fontFamilies : []).map((family) => [family?.name, family]),
  );
  manifest.fontFamilies = visibleNames.map((name) => {
    const family = familyByName.get(name);
    if (!family) throw new Error(`Generated font manifest is missing visible family metadata: ${name}`);
    return family;
  });

  let nextSourceMap = null;
  if (fs.existsSync(sourceMapPath)) {
    const sourceMap = JSON.parse(fs.readFileSync(sourceMapPath, 'utf8'));
    sourceMap.visibleFamilies = visibleNames;
    nextSourceMap = `${JSON.stringify(sourceMap, null, 2)}\n`;
  }

  const allFontsBytes = Buffer.from(nextAllFonts);
  updateAssetMetadata(manifest, ALL_FONTS_PATH, allFontsBytes);
  if (nextSourceMap !== null) {
    updateAssetMetadata(manifest, FONT_SOURCE_MAP_PATH, Buffer.from(nextSourceMap));
  }
  if (Array.isArray(manifest.assets)) {
    manifest.totalBytes = manifest.assets.reduce(
      (total, asset) => total + (Number.isSafeInteger(asset?.bytes) ? asset.bytes : 0),
      0,
    );
  }

  fs.writeFileSync(allFontsPath, nextAllFonts);
  if (nextSourceMap !== null) fs.writeFileSync(sourceMapPath, nextSourceMap);
  fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);
  return { visibleNames };
}

function main() {
  const root = process.argv[2];
  if (!root) throw new Error('Usage: node scripts/apply-font-picker-policy.mjs <font-assets-dir>');
  const result = applyFontPickerPolicy(root);
  process.stdout.write(`Applied curated font picker policy (${result.visibleNames.length} visible families)\n`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  }
}
