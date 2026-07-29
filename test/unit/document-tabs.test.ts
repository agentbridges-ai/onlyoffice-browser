import { describe, expect, it, vi } from 'vitest';
import {
  createDocumentTabId,
  documentTypeForName,
  hasReadPermission,
  requestReadWritePermission,
} from '../../src/pwa/document-tabs';

describe('PWA document tabs', () => {
  it('maps supported file extensions to the matching runtime pack', () => {
    expect(documentTypeForName('notes.docx')).toBe('word');
    expect(documentTypeForName('budget.XLSX')).toBe('cell');
    expect(documentTypeForName('data.csv')).toBe('cell');
    expect(documentTypeForName('deck.pptx')).toBe('slide');
  });

  it('creates distinct stable identifiers without relying on file names', () => {
    expect(createDocumentTabId()).not.toBe(createDocumentTabId());
  });

  it('requests write access only when it is not already granted', async () => {
    const granted = {
      queryPermission: vi.fn().mockResolvedValue('granted'),
      requestPermission: vi.fn(),
    } as unknown as FileSystemFileHandle;
    expect(await requestReadWritePermission(granted)).toBe(true);
    expect(
      (granted as never as { requestPermission: ReturnType<typeof vi.fn> }).requestPermission,
    ).not.toHaveBeenCalled();

    const prompt = {
      queryPermission: vi.fn().mockResolvedValue('prompt'),
      requestPermission: vi.fn().mockResolvedValue('granted'),
    } as unknown as FileSystemFileHandle;
    expect(await requestReadWritePermission(prompt)).toBe(true);
  });

  it('does not treat a persisted handle as readable after permission is revoked', async () => {
    const handle = {
      queryPermission: vi.fn().mockResolvedValue('denied'),
    } as unknown as FileSystemFileHandle;
    expect(await hasReadPermission(handle)).toBe(false);
  });
});
