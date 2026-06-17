import type { DocumentType } from './document-types';

/**
 * Get base path based on deployment environment
 * - GitHub Pages: uses /document/ path
 * - Docker/Other: uses root path /
 */
export const getBasePath = (): string => {
  if (typeof window === 'undefined') {
    return '/';
  }

  const pathname = window.location.pathname;
  // Check if we're in GitHub Pages (path starts with /document/ or contains /document/)
  if (pathname.startsWith('/document/') || pathname === '/document') {
    return '/document/';
  }
  // Docker or other deployments use root path
  return '/';
};

export const BASE_PATH = getBasePath();

/**
 * Get document type from file extension
 */
export function getDocumentType(fileType: string): string | null {
  const type = fileType.toLowerCase();
  if (type === 'docx' || type === 'doc') {
    return 'word';
  } else if (type === 'xlsx' || type === 'xls' || type === 'csv') {
    return 'cell';
  } else if (type === 'pptx' || type === 'ppt') {
    return 'slide';
  }
  return null;
}

export function getMimeTypeFromExtension(extension: string): string {
  const mimeMap: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    svg: 'image/svg+xml',
    tiff: 'image/tiff',
    tif: 'image/tiff',
    emf: 'image/emf',
    wmf: 'image/wmf',
  };
  return mimeMap[extension?.toLowerCase() || ''] || 'image/png';
}

/**
 * Document type mapping
 */
export const DOCUMENT_TYPE_MAP: Record<string, DocumentType> = {
  docx: 'word',
  doc: 'word',
  odt: 'word',
  rtf: 'word',
  txt: 'word',
  xlsx: 'cell',
  xls: 'cell',
  ods: 'cell',
  csv: 'cell',
  pptx: 'slide',
  ppt: 'slide',
  odp: 'slide',
};
