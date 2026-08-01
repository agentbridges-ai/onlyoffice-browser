/**
 * Families shown in the editor font picker.
 *
 * The complete package still carries symbol, math, emoji, historic-script and
 * document-compatibility faces so DocumentServer can preserve existing files
 * and perform its native closest-match fallback. Those support faces must not
 * crowd the user-facing picker.
 */
export const COMMON_FONT_PICKER_FAMILIES = Object.freeze([
  'AR PL UKai CN',
  'AR PL UKai HK',
  'AR PL UKai TW',
  'AR PL UKai TW MBE',
  'Andale Mono',
  'Aptos',
  'Arial',
  'Arial Black',
  'Caladea',
  'Calibri',
  'Cambria',
  'Carlito',
  'Comic Sans MS',
  'Consolas',
  'Courier New',
  'DejaVu Sans',
  'DejaVu Sans Condensed',
  'DejaVu Sans Light',
  'DejaVu Sans Mono',
  'DejaVu Serif',
  'DejaVu Serif Condensed',
  'DengXian',
  'Droid Sans Fallback',
  'FangSong',
  'FreeMono',
  'FreeSans',
  'FreeSerif',
  'Georgia',
  'Impact',
  'KaiTi',
  'Liberation Mono',
  'Liberation Sans',
  'Liberation Sans Narrow',
  'Liberation Serif',
  'Microsoft YaHei',
  'NSimSun',
  'Open Sans',
  'Open Sans Extrabold',
  'Open Sans Light',
  'Open Sans Semibold',
  'SimHei',
  'SimSun',
  'SimSun-ExtB',
  'Times New Roman',
  'Trebuchet MS',
  'Ubuntu',
  'Ubuntu Condensed',
  'Ubuntu Light',
  'Ubuntu Mono',
  'Verdana',
  'WenQuanYi Zen Hei',
  'WenQuanYi Zen Hei Mono',
  'WenQuanYi Zen Hei Sharp',
]);

export const REQUIRED_DEFAULT_FONT_PICKER_FAMILIES = Object.freeze(['Aptos', 'DengXian']);

const commonFontPickerFamilySet = new Set(COMMON_FONT_PICKER_FAMILIES);

export function isCommonFontPickerFamily(name) {
  return typeof name === 'string' && commonFontPickerFamilySet.has(name);
}
