export type OfficeHostBootstrap = {
  sessionId: string;
  parentOrigin: string;
  releaseId: string;
};

const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;

function fragmentParams(url: URL): URLSearchParams {
  return new URLSearchParams(url.hash.startsWith('#') ? url.hash.slice(1) : url.hash);
}

export function releaseIdFromOfficeHostUrl(url: URL): string {
  const match = /^\/r\/([^/]+)\/office-host\.html$/.exec(url.pathname);
  if (!match) return '';
  try {
    const releaseId = decodeURIComponent(match[1]);
    return RELEASE_ID_PATTERN.test(releaseId) ? releaseId : '';
  } catch {
    return '';
  }
}

export function readOfficeHostBootstrap(url: URL): OfficeHostBootstrap {
  const fragment = fragmentParams(url);
  return {
    sessionId: fragment.get('sessionId') || url.searchParams.get('sessionId') || '',
    parentOrigin: fragment.get('parentOrigin') || url.searchParams.get('parentOrigin') || '',
    releaseId: releaseIdFromOfficeHostUrl(url),
  };
}

export function writeOfficeHostBootstrap(
  url: URL,
  bootstrap: Pick<OfficeHostBootstrap, 'sessionId' | 'parentOrigin'>,
): URL {
  const fragment = fragmentParams(url);
  fragment.set('sessionId', bootstrap.sessionId);
  fragment.set('parentOrigin', bootstrap.parentOrigin);
  url.searchParams.delete('sessionId');
  url.searchParams.delete('parentOrigin');
  url.hash = fragment.toString();
  return url;
}

export function resolveOriginBoundWorkerUrl(assetUrl: URL, currentLocation: Location = window.location): URL {
  const workerUrl = new URL(assetUrl.href);
  if (workerUrl.origin === currentLocation.origin) return workerUrl;
  const currentOrigin = new URL(currentLocation.origin);
  workerUrl.protocol = currentOrigin.protocol;
  workerUrl.host = currentOrigin.host;
  return workerUrl;
}
