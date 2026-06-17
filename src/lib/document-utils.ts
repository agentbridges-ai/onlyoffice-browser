import type { DocumentType } from './document-types';

/**
 * Get base path based on deployment environment
 * - GitHub Pages project sites use /<repo>/ paths
 * - Same-origin app deployments usually use root path /
 */
export const getBasePath = (): string => {
  if (typeof window === 'undefined') {
    return '/';
  }

  const pathname = window.location.pathname;
  const projectSiteNames = ['onlyoffice-browser', 'document'];
  const firstSegment = pathname.split('/').filter(Boolean)[0];
  if (firstSegment && projectSiteNames.includes(firstSegment)) {
    return `/${firstSegment}/`;
  }

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
