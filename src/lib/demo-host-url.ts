import type { OfficeHostUrlContext, OfficeHostUrlResolver } from './office-editor';

const PRODUCTION_DEMO_HOST = 'onlyoffice.getpi.work';
const LEGACY_R2_HOST_PATTERN = /^pub-[a-f0-9]+\.r2\.dev$/i;
const LEGACY_STORAGE_KEY = 'onlyoffice-browser:last-host-url';

function productionHostUrl({ sessionId }: OfficeHostUrlContext): string {
  const hostLabel = sessionId.replace(/^office-editor-/, '');
  return `https://${hostLabel}.onlyoffice.getpi.work/office-host.html`;
}

function configuredHostUrlResolver(configured: string, pageUrl: URL): OfficeHostUrlResolver {
  if (configured.includes('{sessionId}')) {
    return ({ sessionId }) =>
      new URL(configured.replaceAll('{sessionId}', sessionId), pageUrl).href;
  }
  return new URL(configured, pageUrl).href;
}

export function resolveDemoHostUrl(pageUrl: URL): OfficeHostUrlResolver {
  const configured = pageUrl.searchParams.get('hostUrl');
  if (pageUrl.hostname === PRODUCTION_DEMO_HOST) {
    if (!configured) return productionHostUrl;
    const configuredUrl = new URL(configured, pageUrl);
    if (
      LEGACY_R2_HOST_PATTERN.test(configuredUrl.hostname) ||
      configuredUrl.hostname === PRODUCTION_DEMO_HOST
    ) {
      return productionHostUrl;
    }
  }

  if (configured) return configuredHostUrlResolver(configured, pageUrl);

  const hostUrl = new URL('/office-host.html', pageUrl);
  if (
    hostUrl.hostname === 'localhost' ||
    (hostUrl.hostname.endsWith('.localhost') && hostUrl.hostname !== 'host.localhost')
  ) {
    hostUrl.hostname = 'host.localhost';
  } else if (hostUrl.hostname === 'host.localhost') {
    hostUrl.hostname = 'app.localhost';
  } else if (hostUrl.hostname === '127.0.0.1') {
    hostUrl.hostname = 'localhost';
  } else if (!hostUrl.hostname.endsWith('.localhost')) {
    hostUrl.hostname = `host.${hostUrl.hostname}`;
  }
  return hostUrl.href;
}

export function clearLegacyDemoHostState(location: Location, storage: Storage): void {
  try {
    storage.removeItem(LEGACY_STORAGE_KEY);
  } catch {
    // Storage can be disabled without affecting the resolver.
  }

  if (location.hostname !== PRODUCTION_DEMO_HOST) return;
  const current = new URL(location.href);
  const configured = current.searchParams.get('hostUrl');
  if (!configured) return;
  const configuredUrl = new URL(configured, current);
  if (
    !LEGACY_R2_HOST_PATTERN.test(configuredUrl.hostname) &&
    configuredUrl.hostname !== PRODUCTION_DEMO_HOST
  ) {
    return;
  }
  current.searchParams.delete('hostUrl');
  window.history.replaceState(window.history.state, '', current);
}
