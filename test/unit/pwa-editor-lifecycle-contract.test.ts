import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';

import { OFFICE_EDITOR_ORIGIN_SLOTS } from '../../src/lib/office-origin-pool';

const source = readFileSync(resolve(process.cwd(), 'src/index.ts'), 'utf8');
const tabStoreSource = readFileSync(resolve(process.cwd(), 'src/pwa/document-tabs.ts'), 'utf8');

describe('standalone PWA editor lifecycle contract', () => {
  it('keeps a mounted editor record for every open document', () => {
    expect(source).toContain('mountOfficeEditor');
    expect(source).toContain('editor?: DocumentEditorRecord;');
    expect(source).toContain('record.mount.activate()');
    expect(source).toContain('const active = candidate === tab;');
    expect(source).toContain("container.setAttribute('inert', '');");
    expect(source).toContain("container.removeAttribute('inert');");
    expect(source).not.toContain('elements.slot.replaceChildren();');
    expect(source).not.toContain('async function destroyEditor()');
  });

  it('keeps the Office surface aligned with the Piwork preview-pane structure', () => {
    expect(source).toContain('data-testid="office-preview-pane"');
    expect(source).toContain('data-testid="office-preview-tabbar"');
    expect(source).toContain('class="document-tabs piwork-scrollbar-hidden"');
    expect(source).toContain('data-testid="office-preview-toolbar"');
    expect(source).toContain("select.className = 'document-tab-select'");
    expect(source).toContain("className = 'document-tab-close'");
    expect(source).toContain('documentTypeSymbol(tab.name)');
    expect(source).toContain('document-type-icon');
  });

  it('bounds open documents to the fixed constellation origin pool', () => {
    expect(OFFICE_EDITOR_ORIGIN_SLOTS).toHaveLength(12);
    expect(source).toContain('const MAX_OPEN_DOCUMENTS = OFFICE_EDITOR_ORIGIN_SLOTS.length;');
    expect(source).toContain('if (tabs.length >= MAX_OPEN_DOCUMENTS)');
    expect(source).toContain('copy.openLimitReached');
    expect(source).toContain('originSlot(tab.editor?.origin)');
    expect(source).toContain('preferredHostSlot: tab.originSlot');
    expect(source).toContain('tab.originSlot = originSlot(record.origin) || undefined;');
    expect(tabStoreSource).toContain('originSlot?: OfficeEditorOriginSlot;');
  });

  it('serializes activation while retaining background documents for instant switching', () => {
    expect(source).toContain('const pendingEditorActivations: DocumentEditorRecord[] = [];');
    expect(source).toContain('while (activeEditorActivations < 1 && pendingEditorActivations.length > 0)');
    expect(source).toContain('record.foreground ||= foreground;');
    expect(source).toContain('if (!tabs.includes(tab) || tab === activeTab || tab.inaccessible || tab.editor) return;');
    expect(source).toContain('await ensureEditor(tab, false);');
    expect(source).toContain('await disposeEditorRecord(tab);');
  });
});
