import { describe, expect, it } from 'vitest';

import {
  isOfficePluginResultForRuntime,
  resolveOfficePluginReady,
  type OfficePluginRuntime,
} from '../../src/lib/office-plugin-runtime';

type TestSource = { name: string };

function runtime(instanceId: string, source: TestSource): OfficePluginRuntime<TestSource> {
  return {
    pluginGuid: 'asc.test',
    pluginInstanceId: instanceId,
    source,
  };
}

describe('Office plugin runtime identity', () => {
  it('deduplicates READY from the same plugin runtime', () => {
    const source = { name: 'plugin-frame' };
    const current = runtime('instance-1', source);

    expect(resolveOfficePluginReady(current, runtime('instance-1', source))).toEqual({
      kind: 'duplicate',
      runtime: current,
    });
  });

  it('replaces the active runtime when a frame reload reuses its WindowProxy', () => {
    const source = { name: 'plugin-frame' };
    const current = runtime('instance-1', source);
    const replacement = runtime('instance-2', source);

    expect(resolveOfficePluginReady(current, replacement)).toEqual({
      kind: 'replaced',
      runtime: replacement,
    });
  });

  it('does not let another window claim an existing runtime identity', () => {
    const current = runtime('instance-1', { name: 'current-frame' });

    expect(resolveOfficePluginReady(current, runtime('instance-1', { name: 'other-frame' }))).toEqual({
      kind: 'ignored',
      runtime: current,
    });
  });

  it('accepts results only from the runtime that received the request', () => {
    const source = { name: 'plugin-frame' };
    const pendingRuntime = runtime('instance-1', source);

    expect(isOfficePluginResultForRuntime(pendingRuntime, 'asc.test', 'instance-1', source)).toBe(true);
    expect(isOfficePluginResultForRuntime(pendingRuntime, 'asc.test', 'instance-2', source)).toBe(false);
    expect(isOfficePluginResultForRuntime(pendingRuntime, 'asc.test', 'instance-1', { name: 'plugin-frame' })).toBe(
      false,
    );
  });
});
