import { Alert, Button, ProgressBar } from '@heroui/react';
import { useSyncExternalStore } from 'react';
import type { OfficeRuntimeResourceManager, OfficeRuntimeResourceSnapshot } from '../lib/runtime-resources';
import type { OfficeCopy } from './i18n';

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(bytes >= 100 * 1024 ** 2 ? 0 : 1)} MB`;
}

function percent(value: number, total: number): number {
  return total > 0 ? Math.min(100, Math.round((value / total) * 100)) : 0;
}

function phaseLabel(snapshot: OfficeRuntimeResourceSnapshot, copy: OfficeCopy): string {
  if (snapshot.phase === 'verifying') return copy.verifying;
  if (snapshot.phase === 'repairing') return copy.repairing;
  if (snapshot.phase === 'paused') return copy.paused;
  if (snapshot.phase === 'activating') return copy.activating;
  if (snapshot.phase === 'planning') return copy.planning;
  if (snapshot.phase === 'downloading') return copy.downloading;
  return copy.readyToDownload;
}

function phaseStep(phase: OfficeRuntimeResourceSnapshot['phase']): number {
  if (phase === 'planning' || phase === 'repairing') return 1;
  if (phase === 'downloading' || phase === 'paused' || phase === 'idle') return 2;
  if (phase === 'verifying') return 3;
  return 4;
}

function phaseProgress(snapshot: OfficeRuntimeResourceSnapshot): { completed: number; total: number } {
  if (snapshot.phase === 'verifying' || snapshot.phase === 'repairing') {
    return { completed: snapshot.verifiedBytes, total: snapshot.verifyBytes };
  }
  if (snapshot.phase === 'activating') {
    const total = Math.max(snapshot.downloadBytes, snapshot.verifyBytes);
    return { completed: total, total };
  }
  return { completed: snapshot.downloadedBytes, total: snapshot.downloadBytes };
}

function errorLabel(snapshot: OfficeRuntimeResourceSnapshot, copy: OfficeCopy): string {
  const code = snapshot.error?.code || snapshot.failedResources[0]?.code;
  return code ? copy.resourceErrors[code] : '';
}

export function OfficeResourcePanel({ manager, copy }: { manager: OfficeRuntimeResourceManager; copy: OfficeCopy }) {
  const snapshot = useSyncExternalStore(
    (listener) => manager.subscribe(listener),
    () => manager.getSnapshot(),
    () => manager.getSnapshot(),
  );
  const busy = snapshot.phase !== 'idle';
  const { completed, total } = phaseProgress(snapshot);
  const progressValue = percent(completed, total);
  const showProgress = busy || (snapshot.readiness !== 'ready' && total > 0);
  const stage = phaseStep(snapshot.phase);
  const segmentLabel =
    snapshot.currentChunkCount > 0 && snapshot.currentChunkIndex > 0
      ? `${copy.packageSegment} ${snapshot.currentChunkIndex}/${snapshot.currentChunkCount}`
      : '';
  const progressValueLabel = `${phaseLabel(snapshot, copy)} · ${formatBytes(completed)} / ${formatBytes(total)}`;

  return (
    <div className="resource-panel">
      <div className="resource-panel__identity">
        <div>
          <strong>OnlyOffice {snapshot.packageVersion}</strong>
          <span>{copy.resourceIntro}</span>
        </div>
        <span data-state={snapshot.readiness}>
          {snapshot.readiness === 'ready'
            ? copy.resourcesReady
            : snapshot.readiness === 'paused'
              ? copy.paused
              : snapshot.readiness === 'error' || snapshot.readiness === 'repair-needed'
                ? copy.resourcesError
                : snapshot.readiness === 'update-available'
                  ? copy.resourceUpdateAvailable
                  : copy.resourcesNeeded}
        </span>
      </div>

      {showProgress && (
        <ProgressBar
          aria-label={copy.resourceProgress}
          className="resource-progress"
          isIndeterminate={snapshot.phase === 'planning'}
          maxValue={Math.max(total, 1)}
          value={Math.min(completed, Math.max(total, 1))}
          valueLabel={progressValueLabel}
        >
          <div className="resource-progress__label">
            <span className="resource-progress__stage">
              <span>
                {copy.resourceStage} {stage}/4
              </span>
              <strong>{phaseLabel(snapshot, copy)}</strong>
              {segmentLabel && <span>{segmentLabel}</span>}
            </span>
            <span>
              {formatBytes(completed)} / {formatBytes(total)}
            </span>
          </div>
          <ProgressBar.Track>
            <ProgressBar.Fill />
          </ProgressBar.Track>
          <ProgressBar.Output>{progressValue}%</ProgressBar.Output>
        </ProgressBar>
      )}

      <div className="resource-package-card" data-ready={snapshot.readiness === 'ready'}>
        <div>
          <strong>{copy.completePackage}</strong>
          <span>{copy.packageIncludes}</span>
        </div>
        <div>
          <strong>{snapshot.readiness === 'ready' ? copy.installed : total > 0 ? formatBytes(total) : '—'}</strong>
          <span>{copy.completePackageHint}</span>
        </div>
      </div>

      {errorLabel(snapshot, copy) && (
        <Alert status="danger">
          <Alert.Indicator />
          <Alert.Content>
            <Alert.Title>{errorLabel(snapshot, copy)}</Alert.Title>
            {snapshot.failedResources[0]?.path && (
              <Alert.Description>{snapshot.failedResources[0].path}</Alert.Description>
            )}
          </Alert.Content>
        </Alert>
      )}

      <div className="resource-actions">
        {snapshot.canResume ? (
          <Button size="sm" variant="primary" onPress={() => void manager.resume()}>
            {copy.resume}
          </Button>
        ) : snapshot.canRetry ? (
          <Button size="sm" variant="primary" onPress={() => void manager.repair({ scope: 'all' })}>
            {copy.retry}
          </Button>
        ) : snapshot.readiness === 'ready' ? (
          <Button
            isDisabled={busy}
            isPending={snapshot.operation === 'check-health'}
            size="sm"
            variant="secondary"
            onPress={() => void manager.repair({ scope: 'all' })}
          >
            {copy.repair}
          </Button>
        ) : (
          <Button
            isDisabled={busy}
            isPending={snapshot.operation === 'load-all'}
            size="sm"
            variant="primary"
            onPress={() => void manager.loadAll()}
          >
            {copy.basicPreset}
          </Button>
        )}
        {snapshot.canPause && (
          <Button size="sm" variant="secondary" onPress={() => manager.pause()}>
            {copy.pause}
          </Button>
        )}
      </div>
    </div>
  );
}
