export const OFFICE_EDITOR_ORIGIN_SLOTS = [
  'aries',
  'taurus',
  'gemini',
  'cancer',
  'leo',
  'virgo',
  'libra',
  'scorpio',
  'sagittarius',
  'capricorn',
  'aquarius',
  'pisces',
] as const;

export type OfficeEditorOriginSlot = (typeof OFFICE_EDITOR_ORIGIN_SLOTS)[number];

const officeEditorOriginSlotSet = new Set<string>(OFFICE_EDITOR_ORIGIN_SLOTS);

export function isOfficeEditorOriginSlot(value: string): value is OfficeEditorOriginSlot {
  return officeEditorOriginSlotSet.has(value);
}

export function isProductionOfficeEditorHostname(hostname: string): boolean {
  const normalized = hostname.toLowerCase();
  if (!normalized.endsWith('.getpi.work')) return false;
  const label = normalized.slice(0, -'.getpi.work'.length);
  return !label.includes('.') && isOfficeEditorOriginSlot(label);
}
