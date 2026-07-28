export const PRODUCTION_OFFICE_HOST_SLOTS = [
  'https://office-misaka.getpi.work/office-host.html',
  'https://office-pectics.getpi.work/office-host.html',
] as const;

export function productionOfficeHostSlots(pageUrl: URL): string[] {
  if (pageUrl.hostname !== 'onlyoffice.getpi.work' || pageUrl.searchParams.has('hostUrl')) return [];
  return [...PRODUCTION_OFFICE_HOST_SLOTS];
}

export class OfficeHostSlotPool {
  private readonly available: string[];
  private readonly leasedOrigins = new Set<string>();

  constructor(hostUrls: string[]) {
    this.available = hostUrls.map((hostUrl) => new URL(hostUrl).href);
  }

  acquire(): string | null {
    const hostUrl = this.available.find((candidate) => !this.leasedOrigins.has(new URL(candidate).origin));
    if (!hostUrl) return null;
    this.leasedOrigins.add(new URL(hostUrl).origin);
    return hostUrl;
  }

  release(hostUrl: string): void {
    this.leasedOrigins.delete(new URL(hostUrl).origin);
  }

  get size(): number {
    return this.available.length;
  }
}
