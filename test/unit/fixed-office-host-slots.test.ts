import { describe, expect, it } from 'vitest';
import {
  OfficeHostSlotPool,
  productionOfficeHostSlots,
} from '../../src/lib/fixed-office-host-slots';

describe('fixed production Office host slots', () => {
  it('uses the two named production slots unless a host override is present', () => {
    expect(productionOfficeHostSlots(new URL('https://onlyoffice.getpi.work/'))).toEqual([
      'https://office-misaka.getpi.work/office-host.html',
      'https://office-pectics.getpi.work/office-host.html',
    ]);
    expect(
      productionOfficeHostSlots(
        new URL('https://onlyoffice.getpi.work/?hostUrl=https://custom.example/office-host.html'),
      ),
    ).toEqual([]);
  });

  it('leases each origin once and makes it reusable after release', () => {
    const pool = new OfficeHostSlotPool([
      'https://office-misaka.getpi.work/office-host.html',
      'https://office-pectics.getpi.work/office-host.html',
    ]);

    const first = pool.acquire();
    const second = pool.acquire();
    expect(first).toBe('https://office-misaka.getpi.work/office-host.html');
    expect(second).toBe('https://office-pectics.getpi.work/office-host.html');
    expect(pool.acquire()).toBeNull();

    pool.release(first!);
    expect(pool.acquire()).toBe(first);
  });
});
