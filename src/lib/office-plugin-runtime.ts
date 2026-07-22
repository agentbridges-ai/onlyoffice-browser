export type OfficePluginRuntime<TSource> = {
  pluginGuid: string;
  pluginInstanceId: string;
  source: TSource;
};

export type OfficePluginReadyDecision<TSource> =
  | { kind: 'registered' | 'replaced'; runtime: OfficePluginRuntime<TSource> }
  | { kind: 'duplicate' | 'ignored'; runtime: OfficePluginRuntime<TSource> };

export function resolveOfficePluginReady<TSource>(
  current: OfficePluginRuntime<TSource> | undefined,
  incoming: OfficePluginRuntime<TSource>,
): OfficePluginReadyDecision<TSource> {
  if (!current) return { kind: 'registered', runtime: incoming };

  if (current.pluginInstanceId === incoming.pluginInstanceId) {
    return current.source === incoming.source
      ? { kind: 'duplicate', runtime: current }
      : { kind: 'ignored', runtime: current };
  }

  return { kind: 'replaced', runtime: incoming };
}

export function isOfficePluginResultForRuntime<TSource>(
  runtime: OfficePluginRuntime<TSource>,
  pluginGuid: string,
  pluginInstanceId: string,
  source: TSource,
): boolean {
  return (
    runtime.pluginGuid === pluginGuid && runtime.pluginInstanceId === pluginInstanceId && runtime.source === source
  );
}
