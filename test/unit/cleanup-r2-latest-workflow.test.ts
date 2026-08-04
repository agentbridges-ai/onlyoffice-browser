import fs from 'node:fs';
import { describe, expect, it } from 'vitest';

const workflow = fs.readFileSync('.github/workflows/cleanup-r2-latest.yml', 'utf8');

describe('latest-only R2 cleanup workflow', () => {
  it('is explicit, dry-run by default, and bound to one stable release', () => {
    expect(workflow).toContain('workflow_dispatch:');
    expect(workflow).toContain("inputs.confirm_cleanup == 'PURGE_LATEST'");
    expect(workflow).toContain('stable_release_id:');
    expect(workflow).toContain('default: true');
    expect(workflow).toContain('if: ${{ !inputs.dry_run }}');
    expect(workflow).toContain('plan-r2-latest-cleanup.mjs');
    expect(workflow).toContain('legacy.manifestUrl === undefined');
    expect(workflow).toContain('legacy.manifestSha256 === undefined');
    expect(workflow).toContain('cmp stable-v5-r2.json stable-v5-public.json');
    expect(workflow).toContain('cmp stable-v5-r2.json stable-v5-after-cleanup.json');
    expect(workflow).toContain('rclone lsjson r2:onlyoffice-getpi-work');
    expect(workflow).toContain('--files-from-raw latest-only-delete-list.txt');
  });

  it('never uses an unbounded bucket sync or delete-after operation', () => {
    expect(workflow).not.toMatch(/rclone\s+sync/);
    expect(workflow).not.toContain('--delete-after');
    expect(workflow).toContain('actions/upload-artifact@');
    expect(workflow).toContain('retention-days: 90');
  });
});
