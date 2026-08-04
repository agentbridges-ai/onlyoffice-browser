import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
const releaseEnvelope: any = await import('../../scripts/' + 'release-envelope.mjs');
const { assertPublicApiEquivalent, assertTarballIntegrity, createReleaseEnvelope, validateReleaseEnvelope } =
  releaseEnvelope;

const temporaryDirectories: string[] = [];
function temp() {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-envelope-'));
  temporaryDirectories.push(directory);
  return directory;
}
afterEach(() =>
  temporaryDirectories.splice(0).forEach((directory) => fs.rmSync(directory, { recursive: true, force: true })),
);

describe('release envelope', () => {
  it('binds immutable runtime, source and optional packed npm identities', () => {
    const workspace = temp();
    const root = path.join(workspace, 'release');
    const releaseId = 'v0.5.12-deadbeef';
    const sourceCommit = 'a'.repeat(40);
    fs.mkdirSync(path.join(root, 'releases', releaseId), { recursive: true });
    fs.mkdirSync(path.join(root, 'channels'), { recursive: true });
    fs.mkdirSync(path.join(workspace, 'dist'), { recursive: true });
    fs.mkdirSync(path.join(workspace, 'dist', 'npm'), { recursive: true });
    fs.writeFileSync(path.join(workspace, 'dist', 'npm', 'public-api.js'), 'export const api = 1;');
    fs.writeFileSync(path.join(workspace, 'dist', 'npm', 'public-api.d.ts'), 'export declare const api: number;');
    const runtimeBytes = Buffer.from('{"version":2}');
    fs.writeFileSync(path.join(workspace, 'dist', 'onlyoffice-runtime-assets.json'), runtimeBytes);
    const manifestBytes = Buffer.from(
      JSON.stringify({
        version: 5,
        releaseId,
        packageVersion: '0.5.12',
        sourceCommit,
        protocolHostBuildId: 'office-host-0.5.12-r1',
        hostBuildId: 'b'.repeat(64),
        runtimeManifestSha256: crypto.createHash('sha256').update(runtimeBytes).digest('hex'),
        x2t: { version: '9.3.0+2', commit: 'c'.repeat(40), sha256: 'd'.repeat(64) },
      }),
    );
    fs.writeFileSync(path.join(root, 'releases', releaseId, 'manifest.json'), manifestBytes);
    fs.writeFileSync(
      path.join(root, 'channels/stable-v5.json'),
      JSON.stringify({
        version: 1,
        releaseId,
        manifestUrl: `/releases/${releaseId}/manifest.json`,
        manifestSha256: crypto.createHash('sha256').update(manifestBytes).digest('hex'),
      }),
    );
    const tarball = path.join(root, 'package.tgz');
    fs.writeFileSync(tarball, 'tarball');
    const envelope = createReleaseEnvelope({ releaseRoot: root, packageTarball: tarball, gitCommit: sourceCommit });
    expect(validateReleaseEnvelope(envelope, { requireRegistry: false })).toMatchObject({
      lifecycle: 'candidate',
      npm: { version: '0.5.12' },
      source: { gitCommit: 'a'.repeat(40) },
      runtime: { releaseId, sourceCommit, protocolHostBuildId: 'office-host-0.5.12-r1' },
    });
    expect(envelope.sourcePackage.integrity).toMatch(/^sha512-/);
    expect(envelope.npm.publicApi).toMatchObject({ jsSha256: expect.stringMatching(/^[a-f0-9]{64}$/) });
    expect(() => validateReleaseEnvelope(envelope)).toThrow(/registry state/);
  });

  it('rejects a release manifest produced by a different candidate commit', () => {
    const workspace = temp();
    const root = path.join(workspace, 'release');
    const releaseId = 'v0.5.12-deadbeef';
    fs.mkdirSync(path.join(root, 'releases', releaseId), { recursive: true });
    fs.mkdirSync(path.join(root, 'channels'), { recursive: true });
    fs.mkdirSync(path.join(workspace, 'dist', 'npm'), { recursive: true });
    fs.writeFileSync(path.join(workspace, 'dist', 'npm', 'public-api.js'), 'api');
    fs.writeFileSync(path.join(workspace, 'dist', 'npm', 'public-api.d.ts'), 'api');
    const runtime = Buffer.from('{}');
    fs.writeFileSync(path.join(workspace, 'dist', 'onlyoffice-runtime-assets.json'), runtime);
    const manifest = Buffer.from(
      JSON.stringify({
        version: 5,
        releaseId,
        packageVersion: '0.5.12',
        sourceCommit: 'b'.repeat(40),
        protocolHostBuildId: 'office-host-0.5.12-r1',
        hostBuildId: 'c'.repeat(64),
        runtimeManifestSha256: crypto.createHash('sha256').update(runtime).digest('hex'),
        x2t: { version: '9.3.0+2', commit: 'd'.repeat(40), sha256: 'e'.repeat(64) },
      }),
    );
    fs.writeFileSync(path.join(root, 'releases', releaseId, 'manifest.json'), manifest);
    fs.writeFileSync(
      path.join(root, 'channels/stable-v5.json'),
      JSON.stringify({
        version: 1,
        releaseId,
        manifestUrl: `/releases/${releaseId}/manifest.json`,
        manifestSha256: crypto.createHash('sha256').update(manifest).digest('hex'),
      }),
    );
    expect(() => createReleaseEnvelope({ releaseRoot: root, gitCommit: 'a'.repeat(40) })).toThrow(/sourceCommit/);
  });

  it('rejects mutable or malformed identities', () => {
    expect(() =>
      validateReleaseEnvelope({
        version: 1,
        lifecycle: 'candidate',
        runtime: { releaseId: '../stable' },
        npm: {},
        source: {},
      }),
    ).toThrow(/releaseId/);
    expect(() =>
      validateReleaseEnvelope({
        version: 1,
        lifecycle: 'candidate',
        runtime: {
          releaseId: 'release',
          manifestUrl: '/releases/release/manifest.json',
          manifestSha256: 'a'.repeat(64),
          runtimeManifestSha256: 'b'.repeat(64),
          hostBuildId: 'c'.repeat(64),
          protocolHostBuildId: 'office-host-0.5.12-r1',
          sourceCommit: 'c'.repeat(40),
        },
        npm: { packageName: '@agentbridges-ai/onlyoffice-browser', version: '0.5.12' },
        source: { gitCommit: 'c'.repeat(64) },
        x2t: { version: '9.3.0+2', commit: 'd'.repeat(40), sha256: 'e'.repeat(64) },
      }),
    ).toThrow(/Git commit/);
  });

  it('requires an already published version to keep exact public API bytes', () => {
    const api = { jsSha256: 'a'.repeat(64), dtsSha256: 'b'.repeat(64) };
    expect(() => assertPublicApiEquivalent(api, { ...api, jsSha256: 'c'.repeat(64) })).toThrow('bump and publish npm');
  });

  it('verifies the downloaded npm tarball before extracting public API files', () => {
    const bytes = Buffer.from('published tgz');
    const integrity = `sha512-${crypto.createHash('sha512').update(bytes).digest('base64')}`;
    expect(assertTarballIntegrity(bytes, integrity)).toBe(integrity);
    expect(() => assertTarballIntegrity(bytes, 'sha512-not-the-tarball')).toThrow(/tarball SHA-512/);
  });
});
