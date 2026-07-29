import { Alert, Button, Disclosure, ProgressBar } from '@heroui/react';
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
  if (snapshot.phase === 'verifying' || snapshot.phase === 'repairing') return copy.verifying;
  if (snapshot.phase === 'paused') return copy.paused;
  if (snapshot.phase === 'activating') return copy.activating;
  if (snapshot.phase === 'planning') return copy.planning;
  return copy.downloading;
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
  const verifying = snapshot.phase === 'verifying' || snapshot.phase === 'repairing';
  const completed = verifying ? snapshot.verifiedBytes : snapshot.downloadedBytes;
  const total = verifying ? snapshot.verifyBytes : snapshot.downloadBytes;
  const progressValue = percent(completed, total);
  const profileLabels = {
    core: copy.core,
    word: copy.word,
    cell: copy.cell,
    slide: copy.slide,
    fonts: copy.fonts,
  };

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

      {(busy || total > 0) && (
        <ProgressBar
          aria-label={copy.resourceProgress}
          className="resource-progress"
          maxValue={Math.max(total, 1)}
          value={Math.min(completed, Math.max(total, 1))}
        >
          <div className="resource-progress__label">
            <span>{phaseLabel(snapshot, copy)}</span>
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

      <div className="resource-profile-grid">
        {snapshot.packs.map((pack) => (
          <div key={pack.id} className="resource-profile-row">
            <span>{profileLabels[pack.id]}</span>
            <span>{pack.ready ? copy.installed : copy.resourcesNeeded}</span>
          </div>
        ))}
      </div>

      {errorLabel(snapshot, copy) && (
        <Alert status="danger">
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
          <Button size="sm" variant="primary" onPress={() => void manager.repair()}>
            {copy.retry}
          </Button>
        ) : (
          <Button
            isDisabled={busy}
            isPending={snapshot.operation === 'prefetch-recommended'}
            size="sm"
            variant="primary"
            onPress={() => void manager.prefetchRecommended()}
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

      <Disclosure className="resource-disclosure">
        <Disclosure.Heading>
          <Disclosure.Trigger>
            {copy.advancedFonts}
            <Disclosure.Indicator />
          </Disclosure.Trigger>
        </Disclosure.Heading>
        <Disclosure.Content>
          <Disclosure.Body>
            <div className="resource-font-list">
              {snapshot.fonts.map((font) => (
                <div key={font.id} className="resource-font-row">
                  <span>
                    {font.name} · {formatBytes(font.bytes)}
                  </span>
                  {font.removable ? (
                    <Button
                      isDisabled={busy}
                      isPending={snapshot.operation === (font.downloaded ? 'remove-font' : 'download-font')}
                      size="sm"
                      variant="secondary"
                      onPress={() =>
                        void (font.downloaded
                          ? manager.uninstallFontFamily(font.id)
                          : manager.downloadFontFamily(font.id))
                      }
                    >
                      {font.downloaded ? copy.remove : copy.download}
                    </Button>
                  ) : (
                    <span>{copy.required}</span>
                  )}
                </div>
              ))}
            </div>
            <div className="resource-actions">
              <Button
                isDisabled={busy}
                isPending={snapshot.operation === 'check-health'}
                size="sm"
                variant="secondary"
                onPress={() => void manager.repair()}
              >
                {copy.repair}
              </Button>
              <Button
                isDisabled={busy}
                isPending={snapshot.operation === 'install-font-preset'}
                size="sm"
                variant="secondary"
                onPress={() => void manager.installFontPreset('office-compatibility')}
              >
                {copy.compatPreset}
              </Button>
              <Button
                isDisabled={busy}
                isPending={snapshot.operation === 'load-all'}
                size="sm"
                variant="secondary"
                onPress={() => void manager.loadAll()}
              >
                {copy.allResources}
              </Button>
            </div>
          </Disclosure.Body>
        </Disclosure.Content>
      </Disclosure>
    </div>
  );
}
