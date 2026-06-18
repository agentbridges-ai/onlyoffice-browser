#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

export const FONT_ASSETS_DIR_ENV = 'ONLYOFFICE_BROWSER_FONT_ASSETS_DIR';
export const GENERATED_FONT_ASSETS_MANIFEST = 'onlyoffice-browser-font-assets.json';

function usage() {
  return `Usage:
  npm run fonts:verify -- --input .onlyoffice-font-assets

Options:
  --input <dir>  Generated OnlyOffice font asset directory. Defaults to ${FONT_ASSETS_DIR_ENV}.
  --help         Show this help.`;
}

function readOption(argv, index, name) {
  const value = argv[index + 1];
  if (!value || value.startsWith('--')) {
    throw new Error(`Missing value for ${name}`);
  }
  return value;
}

export function parseVerifyFontAssetsArgs(argv, env = process.env) {
  const options = {
    input: env[FONT_ASSETS_DIR_ENV] || '',
    help: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    switch (arg) {
      case '--input':
        options.input = readOption(argv, index, arg);
        index += 1;
        break;
      case '--help':
      case '-h':
        options.help = true;
        break;
      default:
        throw new Error(`Unknown option: ${arg}`);
    }
  }

  return options;
}

function assertFile(root, relativePath) {
  const filePath = path.resolve(root, relativePath);
  const relative = path.relative(root, filePath);
  if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error(`Generated font asset path escapes input directory: ${relativePath}`);
  }
  if (!fs.existsSync(filePath) || !fs.statSync(filePath).isFile()) {
    throw new Error(`Generated font asset is missing: ${relativePath}`);
  }
}

function assertStringArray(value, label) {
  if (
    !Array.isArray(value) ||
    value.length === 0 ||
    !value.every((item) => typeof item === 'string' && item.length > 0)
  ) {
    throw new Error(`Generated font asset manifest has invalid ${label}`);
  }
}

export function verifyOnlyOfficeFontAssets(input) {
  if (!input) {
    throw new Error(`Missing required --input <dir> or ${FONT_ASSETS_DIR_ENV}`);
  }

  const root = path.resolve(input);
  if (!fs.existsSync(root) || !fs.statSync(root).isDirectory()) {
    throw new Error(`Generated font asset directory does not exist: ${root}`);
  }

  const manifestPath = path.join(root, GENERATED_FONT_ASSETS_MANIFEST);
  if (!fs.existsSync(manifestPath) || !fs.statSync(manifestPath).isFile()) {
    throw new Error(`Generated font asset manifest is missing: ${GENERATED_FONT_ASSETS_MANIFEST}`);
  }

  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  if (!manifest || manifest.version !== 1) {
    throw new Error(`Generated font asset manifest has unsupported version: ${manifest?.version ?? 'missing'}`);
  }
  if (typeof manifest.allFonts !== 'string' || manifest.allFonts.length === 0) {
    throw new Error('Generated font asset manifest has invalid allFonts');
  }
  if (typeof manifest.fontSelection !== 'string' || manifest.fontSelection.length === 0) {
    throw new Error('Generated font asset manifest has invalid fontSelection');
  }
  assertStringArray(manifest.fontThumbnails, 'fontThumbnails');
  assertStringArray(manifest.fonts, 'fonts');

  assertFile(root, manifest.allFonts);
  assertFile(root, manifest.fontSelection);
  for (const thumbnail of manifest.fontThumbnails) assertFile(root, thumbnail);
  for (const font of manifest.fonts) assertFile(root, font);

  return {
    root,
    fontSet: manifest.fontSet || 'unknown',
    fonts: manifest.fonts.length,
    thumbnails: manifest.fontThumbnails.length,
  };
}

function main() {
  const options = parseVerifyFontAssetsArgs(process.argv.slice(2));
  if (options.help) {
    console.log(usage());
    return;
  }
  const result = verifyOnlyOfficeFontAssets(options.input);
  console.log(
    `Verified OnlyOffice font assets: ${result.root} (${result.fontSet}, ${result.fonts} font files, ${result.thumbnails} thumbnails)`,
  );
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  }
}
