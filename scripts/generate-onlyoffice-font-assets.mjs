#!/usr/bin/env node

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { verifyOnlyOfficeFontAssets } from './verify-onlyoffice-font-assets.mjs';

export const DEFAULT_FONT_GENERATOR_IMAGE = 'onlyoffice/documentserver:9.3.0';
export const FONT_GENERATOR_IMAGE_ENV = 'ONLYOFFICE_BROWSER_FONT_GENERATOR_IMAGE';
export const FONT_SET_ENV = 'ONLYOFFICE_BROWSER_FONT_SET';
export const GENERATED_FONT_ASSETS_MANIFEST = 'onlyoffice-browser-font-assets.json';
export const DEFAULT_FONT_SET = 'zh-core';

const SUPPORTED_FONT_EXTENSIONS = new Set(['.ttf', '.tte', '.otf', '.otc', '.ttc', '.woff', '.woff2']);
const FONT_SETS = new Set(['zh-core', 'full']);
const ZH_CORE_FONT_FAMILIES = [
  'Arial',
  'Aptos',
  'Calibri',
  'Cambria',
  'Cambria Math',
  'Consolas',
  'DengXian',
  'FangSong',
  'KaiTi',
  'Microsoft YaHei',
  'Microsoft YaHei UI',
  'NSimSun',
  'SimHei',
  'SimSun',
  'SimSun-ExtB',
  'Times New Roman',
];
const ZH_CORE_SOURCE_FILE_NAMES = [
  'Aptos-Bold-Italic.ttf',
  'Aptos-Bold.ttf',
  'Aptos-Italic.ttf',
  'Aptos.ttf',
  'Calibri.ttf',
  'Calibrib.ttf',
  'Calibrii.ttf',
  'Calibriz.ttf',
  'Cambria.ttc',
  'Consola.ttf',
  'Consolab.ttf',
  'Deng.ttf',
  'Dengb.ttf',
  'Fangsong.ttf',
  'Kaiti.ttf',
  'SimHei.ttf',
  'Simsun.ttc',
  'arial.ttf',
  'arialbd.ttf',
  'arialbi.ttf',
  'ariali.ttf',
  'msyh.ttc',
  'msyhbd.ttc',
  'simsunb.ttf',
  'times.ttf',
  'timesbd.ttf',
  'timesbi.ttf',
  'timesi.ttf',
];
const LATIN_FALLBACK_FONT_FAMILIES = ['Calibri', 'Arial', 'Carlito', 'Liberation Sans', 'DejaVu Sans', 'Open Sans'];
const CJK_FALLBACK_FONT_FAMILIES = [
  'Microsoft YaHei',
  'SimSun',
  'Noto Sans SC',
  'Noto Sans CJK SC',
  'Noto Sans TC',
  'Noto Sans JP',
  'Noto Sans KR',
  'WenQuanYi Zen Hei',
  'Droid Sans Fallback',
  'AR PL UKai CN',
];
const GENERATED_OUTPUT_PATHS = [
  'fonts',
  'server',
  'sdkjs/common/AllFonts.js',
  GENERATED_FONT_ASSETS_MANIFEST,
  'onlyoffice-browser-font-source-map.json',
];

function usage() {
  return `Usage:
  npm run fonts:generate -- --input /path/to/fonts --output .onlyoffice-font-assets

Options:
  --input <dir>    Directory containing font files. Required.
  --output <dir>   Output directory for generated OnlyOffice font assets. Required.
  --image <image>  Docker image. Defaults to ${DEFAULT_FONT_GENERATOR_IMAGE}.
  --font-set <set> Font set to export: zh-core or full. Defaults to ${DEFAULT_FONT_SET}.
  --keep-font <family>
                  Extra exact font family to keep when --font-set zh-core is used.
                  May be repeated.
  --help           Show this help.`;
}

function readOption(argv, index, name) {
  const value = argv[index + 1];
  if (!value || value.startsWith('--')) {
    throw new Error(`Missing value for ${name}`);
  }
  return value;
}

export function parseGenerateFontAssetsArgs(argv, env = process.env) {
  const options = {
    input: '',
    output: '',
    image: env[FONT_GENERATOR_IMAGE_ENV] || DEFAULT_FONT_GENERATOR_IMAGE,
    fontSet: env[FONT_SET_ENV] || DEFAULT_FONT_SET,
    keepFonts: [],
    help: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    switch (arg) {
      case '--input':
        options.input = readOption(argv, index, arg);
        index += 1;
        break;
      case '--output':
        options.output = readOption(argv, index, arg);
        index += 1;
        break;
      case '--image':
        options.image = readOption(argv, index, arg);
        index += 1;
        break;
      case '--font-set':
        options.fontSet = readOption(argv, index, arg);
        index += 1;
        break;
      case '--keep-font':
        options.keepFonts.push(readOption(argv, index, arg));
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

export function isSupportedFontFileName(fileName) {
  return SUPPORTED_FONT_EXTENSIONS.has(path.extname(fileName).toLowerCase());
}

export function collectFontFiles(inputDir) {
  const fontFiles = [];

  function visit(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const entryPath = path.join(dir, entry.name);
      const stat = fs.statSync(entryPath);
      if (stat.isDirectory()) {
        visit(entryPath);
      } else if (stat.isFile() && isSupportedFontFileName(entry.name)) {
        fontFiles.push(entryPath);
      }
    }
  }

  visit(inputDir);
  return fontFiles.sort();
}

function assertDirectory(dir, label) {
  let stat;
  try {
    stat = fs.statSync(dir);
  } catch {
    throw new Error(`${label} does not exist: ${dir}`);
  }

  if (!stat.isDirectory()) {
    throw new Error(`${label} must be a directory: ${dir}`);
  }
}

export function validateGenerateFontAssetsOptions(options) {
  if (options.help) return;
  if (!options.input) throw new Error('Missing required --input <dir>');
  if (!options.output) throw new Error('Missing required --output <dir>');
  if (!options.image) throw new Error('Docker image is empty');
  if (!FONT_SETS.has(options.fontSet)) {
    throw new Error(`Invalid --font-set ${options.fontSet}. Expected one of: ${Array.from(FONT_SETS).join(', ')}`);
  }

  const input = path.resolve(options.input);
  assertDirectory(input, 'Input font directory');

  const fontFiles = collectFontFiles(input);
  if (fontFiles.length === 0) {
    throw new Error(
      `No supported font files found in ${input}. Supported extensions: ${Array.from(SUPPORTED_FONT_EXTENSIONS)
        .sort()
        .join(', ')}`,
    );
  }

  const output = path.resolve(options.output);
  if (fs.existsSync(output) && !fs.statSync(output).isDirectory()) {
    throw new Error(`Output path exists and is not a directory: ${output}`);
  }

  return { input, output, image: options.image, fontSet: options.fontSet, keepFonts: options.keepFonts, fontFiles };
}

function removeIfExists(targetPath) {
  fs.rmSync(targetPath, { recursive: true, force: true });
}

function makeWritableRecursive(targetPath) {
  if (!fs.existsSync(targetPath)) return;
  const stat = fs.statSync(targetPath);
  fs.chmodSync(targetPath, stat.isDirectory() ? 0o755 : 0o644);
  if (!stat.isDirectory()) return;

  for (const entry of fs.readdirSync(targetPath)) {
    makeWritableRecursive(path.join(targetPath, entry));
  }
}

function prepareOutputDirectory(output) {
  fs.mkdirSync(output, { recursive: true });
  makeWritableRecursive(output);

  for (const relativePath of GENERATED_OUTPUT_PATHS) {
    removeIfExists(path.join(output, relativePath));
  }

  const imagesDir = path.join(output, 'sdkjs/common/Images');
  if (fs.existsSync(imagesDir)) {
    for (const entry of fs.readdirSync(imagesDir)) {
      if (/^fonts_thumbnail(?:_ea)?(?:@[\d.]+x)?\.png$/.test(entry)) {
        removeIfExists(path.join(imagesDir, entry));
      }
    }
  }
}

function runCommand(command, args, options = {}) {
  const result = spawnSync(command, args, {
    encoding: 'utf8',
    stdio: options.stdio || 'pipe',
  });

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    const details = [result.stdout, result.stderr].filter(Boolean).join('\n').trim();
    throw new Error(
      `${command} ${args.join(' ')} failed with exit code ${result.status}${details ? `\n${details}` : ''}`,
    );
  }

  return result;
}

function assertDockerAvailable() {
  try {
    runCommand('docker', ['--version']);
  } catch (error) {
    throw new Error(`Docker is required to generate official OnlyOffice font assets: ${error.message}`);
  }
}

function hostUserEnv() {
  const uid = typeof process.getuid === 'function' ? String(process.getuid()) : '';
  const gid = typeof process.getgid === 'function' ? String(process.getgid()) : '';
  return { uid, gid };
}

export function dockerGenerationScript(options) {
  const zhCoreFontFamilies = Array.from(new Set(ZH_CORE_FONT_FAMILIES)).sort();
  const zhCoreSourceFileNames = Array.from(
    new Set(ZH_CORE_SOURCE_FILE_NAMES.map((fileName) => fileName.toLowerCase())),
  ).sort();
  const latinFallbackFontFamilies = Array.from(new Set(LATIN_FALLBACK_FONT_FAMILIES));
  const cjkFallbackFontFamilies = Array.from(new Set(CJK_FALLBACK_FONT_FAMILIES));
  const keepFontFamilies = Array.from(new Set(options.keepFonts)).sort();
  return `
set -euo pipefail

DS_DIR="/var/www/onlyoffice/documentserver"
IN="/onlyoffice-browser-input-fonts"
OUT="/onlyoffice-browser-output-assets"
EXTRA="/usr/share/fonts/onlyoffice-browser-extra"

mkdir -p "$EXTRA"
find "$IN" -type f \\( \\
  -iname '*.ttf' -o -iname '*.tte' -o -iname '*.otf' -o -iname '*.otc' -o \\
  -iname '*.ttc' -o -iname '*.woff' -o -iname '*.woff2' \\
\\) -print0 | xargs -0 -I{} cp -f "{}" "$EXTRA/"

fc-cache -f "$EXTRA" || true
/usr/bin/documentserver-generate-allfonts.sh true true

mkdir -p "$OUT/sdkjs/common/Images" "$OUT/fonts" "$OUT/server/FileConverter/bin"
cp -f "$DS_DIR/sdkjs/common/AllFonts.js" "$OUT/sdkjs/common/AllFonts.js"
find "$DS_DIR/sdkjs/common/Images" -maxdepth 1 -type f \\( \\
  -name 'fonts_thumbnail*.png' -o -name 'fonts_thumbnail_ea*.png' \\
\\) -exec cp -f {} "$OUT/sdkjs/common/Images/" \\;
cp -f "$DS_DIR/server/FileConverter/bin/font_selection.bin" "$OUT/server/FileConverter/bin/font_selection.bin"
if [ -f "$DS_DIR/server/FileConverter/bin/AllFonts.js" ]; then
  cp -f "$DS_DIR/server/FileConverter/bin/AllFonts.js" "$OUT/server/FileConverter/bin/AllFonts.js"
fi

OUT_DIR="$OUT" FONT_SET=${JSON.stringify(options.fontSet)} ZH_CORE_FONT_FAMILIES=${JSON.stringify(JSON.stringify(zhCoreFontFamilies))} ZH_CORE_SOURCE_FILE_NAMES=${JSON.stringify(JSON.stringify(zhCoreSourceFileNames))} LATIN_FALLBACK_FONT_FAMILIES=${JSON.stringify(JSON.stringify(latinFallbackFontFamilies))} CJK_FALLBACK_FONT_FAMILIES=${JSON.stringify(JSON.stringify(cjkFallbackFontFamilies))} KEEP_FONT_FAMILIES=${JSON.stringify(JSON.stringify(keepFontFamilies))} python3 - <<'PY'
import json
import os
import re
import shutil

out = os.environ["OUT_DIR"]
font_set = os.environ["FONT_SET"]
zh_core_font_families = set(json.loads(os.environ["ZH_CORE_FONT_FAMILIES"]))
zh_core_source_file_names = set(json.loads(os.environ["ZH_CORE_SOURCE_FILE_NAMES"]))
latin_fallback_font_families = json.loads(os.environ["LATIN_FALLBACK_FONT_FAMILIES"])
cjk_fallback_font_families = json.loads(os.environ["CJK_FALLBACK_FONT_FAMILIES"])
keep_font_families = set(json.loads(os.environ["KEEP_FONT_FAMILIES"]))
server_allfonts = os.path.join(out, "server/FileConverter/bin/AllFonts.js")
web_allfonts = os.path.join(out, "sdkjs/common/AllFonts.js")
fonts_out = os.path.join(out, "fonts")

def read_source(file_path):
    with open(file_path, "r", encoding="utf-8-sig") as handle:
        return handle.read()

def parse_js_array(source, name):
    match = re.search(r'window\\["' + re.escape(name) + r'"\\]\\s*=\\s*(\\[[\\s\\S]*?\\]);', source)
    if not match:
        raise SystemExit(f"Unable to locate {name}")
    return json.loads(match.group(1))

def replace_js_array(source, name, value):
    replacement = 'window["' + name + '"] = ' + json.dumps(value, ensure_ascii=False, indent=0) + ';'
    return re.sub(r'window\\["' + re.escape(name) + r'"\\]\\s*=\\s*\\[[\\s\\S]*?\\];', replacement, source, count=1)

def upsert_js_array(source, name, value):
    replacement = 'window["' + name + '"] = ' + json.dumps(value, ensure_ascii=False, indent=0) + ';'
    pattern = r'window\\["' + re.escape(name) + r'"\\]\\s*=\\s*\\[[\\s\\S]*?\\];'
    if re.search(pattern, source):
        return re.sub(pattern, replacement, source, count=1)
    marker = 'window["__fonts_infos"]'
    index = source.find(marker)
    if index < 0:
        return replacement + "\\n" + source
    return source[:index] + replacement + "\\n" + source[index:]

def is_cjk_family_name(name):
    lowered = name.lower()
    cjk_markers = [
        "fang", "hei", "jheng", "kai", "ming", "simsun", "song", "ukai", "wqy", "yahei",
        "deng", "gothic", "mincho", "dotum", "gulim", "batang", "gungsuh",
    ]
    return any(marker in lowered for marker in cjk_markers)

def find_font_info(name):
    for info in web_infos:
        if info and info[0] == name:
            return info
    return None

def first_source_index(info):
    if not info:
        return -1
    for slot_index in range(1, len(info), 2):
        source_index = info[slot_index]
        if source_index >= 0:
            return source_index
    return -1

def first_available_source(*family_names):
    for family_name in family_names:
        source_index = first_source_index(find_font_info(family_name))
        if source_index >= 0:
            return family_name, source_index
    return "", -1

def source_file_name(source_index):
    if source_index < 0 or source_index >= len(source_files):
        return ""
    return os.path.basename(source_files[source_index]).lower()

def allowed_same_family_source_index(info, slot_index):
    if info[0] in keep_font_families:
        return info[slot_index]

    slot_priority = {
        1: [1, 5, 3, 7],
        3: [1, 3, 5, 7],
        5: [5, 1, 7, 3],
        7: [5, 1, 7, 3],
    }
    for candidate_slot in slot_priority.get(slot_index, [1, 5, 3, 7]):
        candidate_index = info[candidate_slot]
        if candidate_index >= 0 and source_file_name(candidate_index) in zh_core_source_file_names:
            return candidate_index
    return -1

server_source = read_source(server_allfonts)
web_source = read_source(web_allfonts)
source_files = parse_js_array(server_source, "__fonts_files")
web_infos = parse_js_array(web_source, "__fonts_infos")
web_ranges = parse_js_array(web_source, "__fonts_ranges")

latin_fallback_family_name, latin_fallback_source_index = first_available_source(*latin_fallback_font_families)
cjk_fallback_family_name, cjk_fallback_source_index = first_available_source(*cjk_fallback_font_families)
if latin_fallback_source_index < 0 or cjk_fallback_source_index < 0:
    raise SystemExit(
        "Unable to locate fallback fonts in generated AllFonts.js. "
        + "Latin candidates: "
        + ", ".join(latin_fallback_font_families)
        + "; CJK candidates: "
        + ", ".join(cjk_fallback_font_families)
    )

if font_set == "zh-core":
    kept_family_names = {info[0] for info in web_infos if info and (info[0] in zh_core_font_families or info[0] in keep_font_families)}
    kept_family_names.update(name for name in [latin_fallback_family_name, cjk_fallback_family_name] if name)
    if not kept_family_names:
        raise SystemExit("zh-core font set did not match any generated font families")
else:
    kept_family_names = {info[0] for info in web_infos if info}

used_source_indexes = []
used_source_index_set = set()
new_infos = []
for original_info in web_infos:
    info = list(original_info)
    keep_actual_font = font_set == "full" or info[0] in kept_family_names
    for slot_index in range(1, len(info), 2):
        source_index = info[slot_index]
        if source_index < 0:
            continue
        if font_set == "zh-core" and keep_actual_font and info[0] not in keep_font_families:
            if source_file_name(source_index) not in zh_core_source_file_names:
                same_family_source_index = allowed_same_family_source_index(info, slot_index)
                source_index = same_family_source_index if same_family_source_index >= 0 else (
                    cjk_fallback_source_index if is_cjk_family_name(info[0]) else latin_fallback_source_index
                )
                info[slot_index] = source_index
        elif not keep_actual_font:
            source_index = cjk_fallback_source_index if is_cjk_family_name(info[0]) else latin_fallback_source_index
            info[slot_index] = source_index
        if source_index not in used_source_index_set:
            used_source_index_set.add(source_index)
            used_source_indexes.append(source_index)
    new_infos.append(info)

source_index_map = {old_index: new_index for new_index, old_index in enumerate(used_source_indexes)}
web_files = []
source_map = []
for new_index, old_index in enumerate(used_source_indexes):
    source_path = source_files[old_index]
    if not os.path.isfile(source_path):
        raise SystemExit(f"Font source does not exist: {source_path}")
    _, extension = os.path.splitext(source_path)
    file_name = f"{new_index:03d}{extension.lower() or '.font'}"
    target_path = os.path.join(fonts_out, file_name)
    shutil.copyfile(source_path, target_path)
    web_files.append(file_name)
    source_map.append({"originalIndex": old_index, "index": new_index, "source": source_path, "file": f"fonts/{file_name}"})

for info in new_infos:
    for slot_index in range(1, len(info), 2):
        source_index = info[slot_index]
        if source_index >= 0:
            info[slot_index] = source_index_map[source_index]

web_source = replace_js_array(web_source, "__fonts_files", web_files)
web_source = replace_js_array(web_source, "__fonts_infos", new_infos)
web_source = replace_js_array(web_source, "__fonts_ranges", web_ranges)
if font_set == "zh-core":
    web_source = upsert_js_array(web_source, "__fonts_visible_names", sorted(kept_family_names))

with open(web_allfonts, "w", encoding="utf-8") as handle:
    handle.write(web_source)

with open(os.path.join(out, "onlyoffice-browser-font-source-map.json"), "w", encoding="utf-8") as handle:
    json.dump({
        "fontSet": font_set,
        "keptFamilies": sorted(kept_family_names),
        "listedFamilies": [info[0] for info in new_infos],
        "fonts": source_map,
    }, handle, ensure_ascii=False, indent=2)
    handle.write("\\n")
PY

if [ -n "\${HOST_UID:-}" ] && [ -n "\${HOST_GID:-}" ]; then
  chown -R "$HOST_UID:$HOST_GID" "$OUT" || true
fi
chmod -R u+rwX "$OUT" || true
`;
}

function runDockerGenerator(options) {
  const { uid, gid } = hostUserEnv();
  const containerName = `onlyoffice-browser-fonts-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const args = [
    'run',
    '--rm',
    '--name',
    containerName,
    '-e',
    `HOST_UID=${uid}`,
    '-e',
    `HOST_GID=${gid}`,
    '-v',
    `${options.input}:/onlyoffice-browser-input-fonts:ro`,
    '-v',
    `${options.output}:/onlyoffice-browser-output-assets`,
    '--entrypoint',
    'bash',
    options.image,
    '-lc',
    dockerGenerationScript(options),
  ];

  runCommand('docker', args, { stdio: 'inherit' });
}

function createFontStagingDirectory(fontFiles) {
  const stagingDir = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-browser-font-input-'));
  const usedNames = new Set();

  for (const fontFile of fontFiles) {
    const parsed = path.parse(fontFile);
    let fileName = parsed.base;
    let counter = 1;
    while (usedNames.has(fileName)) {
      fileName = `${parsed.name}-${counter}${parsed.ext}`;
      counter += 1;
    }
    usedNames.add(fileName);
    fs.copyFileSync(fontFile, path.join(stagingDir, fileName));
  }

  return stagingDir;
}

function listGeneratedFontAssetPaths(output) {
  const fontsRoot = path.join(output, 'fonts');
  if (!fs.existsSync(fontsRoot)) return [];

  const result = [];
  function visit(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const entryPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        visit(entryPath);
      } else if (entry.isFile()) {
        result.push(`fonts/${path.relative(fontsRoot, entryPath).replaceAll(path.sep, '/')}`);
      }
    }
  }

  visit(fontsRoot);
  return result.sort();
}

function listGeneratedFontThumbnailPaths(output) {
  const imagesRoot = path.join(output, 'sdkjs/common/Images');
  if (!fs.existsSync(imagesRoot)) return [];

  return fs
    .readdirSync(imagesRoot)
    .filter((entry) => /^fonts_thumbnail(?:_ea)?(?:@[\d.]+x)?\.png$/.test(entry))
    .sort()
    .map((entry) => `sdkjs/common/Images/${entry}`);
}

function assertGeneratedAssets(output) {
  const requiredFiles = ['sdkjs/common/AllFonts.js', 'server/FileConverter/bin/font_selection.bin'];

  for (const relativePath of requiredFiles) {
    const assetPath = path.join(output, relativePath);
    if (!fs.existsSync(assetPath) || !fs.statSync(assetPath).isFile()) {
      throw new Error(`Generated asset is missing: ${relativePath}`);
    }
  }

  if (listGeneratedFontThumbnailPaths(output).length === 0) {
    throw new Error('Generated font thumbnail images are missing');
  }

  if (listGeneratedFontAssetPaths(output).length === 0) {
    throw new Error('Generated web font files are missing');
  }
}

function writeGeneratedManifest(output, options) {
  const manifest = {
    version: 1,
    generator: 'documentserver-generate-allfonts.sh',
    image: options.image,
    fontSet: options.fontSet,
    generatedAt: new Date().toISOString(),
    allFonts: 'sdkjs/common/AllFonts.js',
    fontSelection: 'server/FileConverter/bin/font_selection.bin',
    fontThumbnails: listGeneratedFontThumbnailPaths(output),
    fonts: listGeneratedFontAssetPaths(output),
  };

  fs.writeFileSync(path.join(output, GENERATED_FONT_ASSETS_MANIFEST), `${JSON.stringify(manifest, null, 2)}\n`);
}

export function generateOnlyOfficeFontAssets(options) {
  const validated = validateGenerateFontAssetsOptions(options);
  if (!validated) return;

  console.log(`Using Docker image: ${validated.image}`);
  console.log(`Found ${validated.fontFiles.length} font files in ${validated.input}`);
  assertDockerAvailable();
  const stagingDir = createFontStagingDirectory(validated.fontFiles);
  console.log(`Prepared temporary font input: ${stagingDir}`);
  try {
    prepareOutputDirectory(validated.output);
    console.log(`Prepared output directory: ${validated.output}`);
    console.log('Running official OnlyOffice font generator...');
    runDockerGenerator({ ...validated, input: stagingDir });
    assertGeneratedAssets(validated.output);
    writeGeneratedManifest(validated.output, validated);
    verifyOnlyOfficeFontAssets(validated.output);
  } finally {
    fs.rmSync(stagingDir, { recursive: true, force: true });
  }
}

function main() {
  const options = parseGenerateFontAssetsArgs(process.argv.slice(2));
  if (options.help) {
    console.log(usage());
    return;
  }

  generateOnlyOfficeFontAssets(options);
  console.log(`Generated OnlyOffice font assets at ${path.resolve(options.output)}`);
}

function isDirectRun() {
  if (!process.argv[1]) return false;

  try {
    return fs.realpathSync(process.argv[1]) === fs.realpathSync(fileURLToPath(import.meta.url));
  } catch {
    return import.meta.url === pathToFileURL(process.argv[1]).href;
  }
}

if (isDirectRun()) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : error);
    process.exitCode = 1;
  }
}
