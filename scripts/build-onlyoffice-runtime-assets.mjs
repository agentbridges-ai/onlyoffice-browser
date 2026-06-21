#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DEFAULT_TYPES = ['word', 'cell', 'slide'];
const DEFAULT_DICTIONARIES = ['en_US'];
const PACKS = ['core', ...DEFAULT_TYPES];

const TYPE_CONFIG = {
  word: {
    editorDir: 'web-apps/apps/documenteditor/',
    sdkDir: 'sdkjs/word/',
  },
  cell: {
    editorDir: 'web-apps/apps/spreadsheeteditor/',
    sdkDir: 'sdkjs/cell/',
  },
  slide: {
    editorDir: 'web-apps/apps/presentationeditor/',
    sdkDir: 'sdkjs/slide/',
  },
};

function usage() {
  return `Usage:
  node scripts/build-onlyoffice-runtime-assets.mjs --input public --output .onlyoffice-runtime-assets
  node scripts/build-onlyoffice-runtime-assets.mjs --input dist --prune-root --split-output dist/asset-packs

Options:
  --input <dir>             Runtime asset source or built dist directory.
  --output <dir>            Copy optimized runtime assets to this directory.
  --split-output <dir>      Copy canonical path packs to core/word/cell/slide subdirectories.
  --prune-root              Remove non-selected runtime files from --input in place.
  --types <list>            Comma list: word,cell,slide. Default: word,cell,slide.
  --dictionaries <list>     Comma list of dictionary locales. Default: en_US.
  --keep-help               Keep bundled help resources. Default: remove help resources.
  --dry-run                 Print summary without copying or deleting.
`;
}

function parseList(value) {
  return value
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

export function parseRuntimeAssetArgs(argv = process.argv.slice(2)) {
  const options = {
    input: '',
    output: '',
    splitOutput: '',
    pruneRoot: false,
    types: DEFAULT_TYPES,
    dictionaries: DEFAULT_DICTIONARIES,
    keepHelp: false,
    dryRun: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const next = () => {
      const value = argv[index + 1];
      if (!value || value.startsWith('--')) {
        throw new Error(`${arg} requires a value`);
      }
      index += 1;
      return value;
    };

    if (arg === '--input') {
      options.input = next();
    } else if (arg === '--output') {
      options.output = next();
    } else if (arg === '--split-output') {
      options.splitOutput = next();
    } else if (arg === '--prune-root') {
      options.pruneRoot = true;
    } else if (arg === '--types') {
      options.types = parseList(next());
    } else if (arg === '--dictionaries') {
      options.dictionaries = parseList(next());
    } else if (arg === '--keep-help') {
      options.keepHelp = true;
    } else if (arg === '--dry-run') {
      options.dryRun = true;
    } else if (arg === '--help' || arg === '-h') {
      options.help = true;
    } else {
      throw new Error(`Unknown option: ${arg}`);
    }
  }

  return options;
}

function assertKnownValues(options) {
  const invalidTypes = options.types.filter((type) => !DEFAULT_TYPES.includes(type));
  if (invalidTypes.length > 0) {
    throw new Error(`Invalid --types value: ${invalidTypes.join(', ')}`);
  }
  if (options.types.length === 0) {
    throw new Error('--types must include at least one document type');
  }
  if (options.dictionaries.some((locale) => locale.includes('/') || locale.includes('\\') || locale.includes('..'))) {
    throw new Error('--dictionaries must contain locale directory names only');
  }
}

export function validateRuntimeAssetOptions(options) {
  if (options.help) return options;
  if (!options.input) {
    throw new Error('--input is required');
  }
  if (!options.output && !options.splitOutput && !options.pruneRoot) {
    throw new Error('At least one of --output, --split-output, or --prune-root is required');
  }

  const input = path.resolve(options.input);
  if (!fs.existsSync(input) || !fs.statSync(input).isDirectory()) {
    throw new Error(`Input directory does not exist: ${input}`);
  }

  assertKnownValues(options);

  return {
    ...options,
    input,
    output: options.output ? path.resolve(options.output) : '',
    splitOutput: options.splitOutput ? path.resolve(options.splitOutput) : '',
    types: [...new Set(options.types)],
    dictionaries: [...new Set(options.dictionaries)],
  };
}

function normalizeRelativePath(relativePath) {
  return relativePath.split(path.sep).join('/').replace(/^\/+/, '');
}

function isRuntimeAsset(relativePath) {
  return (
    relativePath.startsWith('web-apps/') ||
    relativePath.startsWith('sdkjs/') ||
    relativePath.startsWith('wasm/') ||
    relativePath.startsWith('libs/') ||
    relativePath.startsWith('dictionaries/') ||
    relativePath.startsWith('fonts/') ||
    relativePath.startsWith('server/FileConverter/') ||
    [
      'document_editor_service_worker.js',
      'plugins.json',
      'reset.html',
      'sw.js',
      'themes.json',
      'onlyoffice-browser-font-assets.json',
    ].includes(relativePath)
  );
}

function isHelpAsset(relativePath) {
  return /\/resources\/help\//.test(relativePath);
}

function dictionaryPack(relativePath, dictionaries) {
  if (!relativePath.startsWith('dictionaries/')) return null;
  const locale = relativePath.slice('dictionaries/'.length).split('/')[0];
  return dictionaries.includes(locale) ? 'core' : null;
}

export function getRuntimeAssetPack(relativePath, options = {}) {
  const normalized = normalizeRelativePath(relativePath);
  const types = options.types || DEFAULT_TYPES;
  const dictionaries = options.dictionaries || DEFAULT_DICTIONARIES;
  const keepHelp = Boolean(options.keepHelp);

  if (!isRuntimeAsset(normalized)) return null;
  if (!keepHelp && isHelpAsset(normalized)) return null;
  if (normalized.startsWith('fonts/')) return null;
  if (normalized.startsWith('server/FileConverter/')) return null;
  if (normalized.startsWith('sdkjs/pdf/') || normalized.startsWith('sdkjs/visio/')) return null;

  const dictionary = dictionaryPack(normalized, dictionaries);
  if (normalized.startsWith('dictionaries/')) return dictionary;

  if (
    normalized.startsWith('web-apps/apps/api/') ||
    normalized.startsWith('web-apps/apps/common/') ||
    normalized.startsWith('web-apps/vendor/') ||
    normalized.startsWith('sdkjs/common/') ||
    normalized.startsWith('wasm/x2t/') ||
    normalized.startsWith('libs/') ||
    [
      'document_editor_service_worker.js',
      'plugins.json',
      'reset.html',
      'sw.js',
      'themes.json',
      'onlyoffice-browser-font-assets.json',
    ].includes(normalized)
  ) {
    return 'core';
  }

  for (const type of types) {
    const config = TYPE_CONFIG[type];
    if (normalized.startsWith(config.editorDir) || normalized.startsWith(config.sdkDir)) {
      return type;
    }
  }

  return null;
}

function walkFiles(root) {
  const files = [];
  const stack = [''];
  while (stack.length > 0) {
    const relative = stack.pop();
    const absolute = path.join(root, relative);
    for (const entry of fs.readdirSync(absolute, { withFileTypes: true })) {
      const entryRelative = path.join(relative, entry.name);
      const entryAbsolute = path.join(root, entryRelative);
      if (entry.isDirectory()) {
        stack.push(entryRelative);
      } else if (entry.isFile()) {
        files.push(normalizeRelativePath(path.relative(root, entryAbsolute)));
      }
    }
  }
  return files;
}

function copyFile(sourceRoot, targetRoot, relativePath) {
  const source = path.join(sourceRoot, relativePath);
  const target = path.join(targetRoot, relativePath);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.copyFileSync(source, target);
}

function removeEmptyDirectories(root) {
  if (!fs.existsSync(root)) return;
  for (const entry of fs.readdirSync(root, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue;
    const absolute = path.join(root, entry.name);
    removeEmptyDirectories(absolute);
    if (fs.existsSync(absolute) && fs.readdirSync(absolute).length === 0) {
      fs.rmdirSync(absolute);
    }
  }
}

function writeManifest(root, manifest) {
  fs.mkdirSync(root, { recursive: true });
  fs.writeFileSync(path.join(root, 'onlyoffice-runtime-assets.json'), `${JSON.stringify(manifest, null, 2)}\n`);
}

export function collectRuntimeAssets(input, options) {
  const selected = [];
  const excluded = [];
  for (const relativePath of walkFiles(input)) {
    const pack = getRuntimeAssetPack(relativePath, options);
    if (pack) {
      selected.push({ path: relativePath, pack });
    } else if (isRuntimeAsset(relativePath)) {
      excluded.push(relativePath);
    }
  }
  return { selected, excluded };
}

export function buildRuntimeAssets(options) {
  const { selected, excluded } = collectRuntimeAssets(options.input, options);
  const manifest = {
    version: 1,
    generatedAt: new Date().toISOString(),
    types: options.types,
    dictionaries: options.dictionaries,
    keepHelp: options.keepHelp,
    packs: Object.fromEntries(PACKS.map((pack) => [pack, selected.filter((asset) => asset.pack === pack).length])),
    selected: selected.length,
    excluded: excluded.length,
  };

  if (options.dryRun) {
    return manifest;
  }

  if (options.output) {
    fs.rmSync(options.output, { recursive: true, force: true });
    for (const asset of selected) {
      copyFile(options.input, options.output, asset.path);
    }
    writeManifest(options.output, manifest);
  }

  if (options.splitOutput) {
    fs.rmSync(options.splitOutput, { recursive: true, force: true });
    for (const asset of selected) {
      copyFile(options.input, path.join(options.splitOutput, asset.pack), asset.path);
    }
    writeManifest(options.splitOutput, manifest);
  }

  if (options.pruneRoot) {
    for (const relativePath of excluded) {
      fs.rmSync(path.join(options.input, relativePath), { force: true });
    }
    removeEmptyDirectories(options.input);
    writeManifest(options.input, manifest);
  }

  return manifest;
}

function main() {
  const options = parseRuntimeAssetArgs();
  if (options.help) {
    console.log(usage());
    return;
  }
  const validated = validateRuntimeAssetOptions(options);
  const manifest = buildRuntimeAssets(validated);
  console.log(
    `Optimized OnlyOffice runtime assets: selected ${manifest.selected}, excluded ${manifest.excluded}, packs ${JSON.stringify(manifest.packs)}`,
  );
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
