#!/usr/bin/env node

// A release envelope is deliberately small and portable.  It is the hand-off
// between the candidate build and the separately approved promotion workflow;
// neither workflow is allowed to infer identity from a mutable channel pointer.
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import { assertNpmSlsaAttestation, assertPublishedIdentity } from './npm-publication-identity.mjs';

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function readJson(file, label) {
  try {
    return { bytes: fs.readFileSync(file), value: JSON.parse(fs.readFileSync(file, 'utf8')) };
  } catch (error) {
    throw new Error(`Cannot read ${label} at ${file}`, { cause: error });
  }
}

function requireDigest(value, label) {
  if (!/^[a-f0-9]{64}$/.test(value || '')) throw new Error(`${label} must be a SHA-256 hex digest`);
  return value;
}

function requireGitCommit(value, label) {
  if (!/^[a-f0-9]{40}$/.test(value || '')) throw new Error(`${label} must be a 40-character Git commit`);
  return value;
}

function requireReleaseId(value) {
  if (!/^[a-zA-Z0-9._+-]{1,128}$/.test(value || '')) throw new Error('releaseId is invalid');
  return value;
}

function requireProtocolHostBuildId(value, packageVersion) {
  const escapedVersion = String(packageVersion).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (!new RegExp(`^office-host-${escapedVersion}-r[1-9][0-9]*$`).test(value || '')) {
    throw new Error('protocol Host build ID is invalid');
  }
  return value;
}

function publicApiIdentity(distRoot) {
  const read = (name) => {
    const file = path.join(distRoot, 'npm', name);
    if (!fs.existsSync(file)) throw new Error(`missing public npm API artifact ${file}`);
    return sha256(fs.readFileSync(file));
  };
  return { jsSha256: read('public-api.js'), dtsSha256: read('public-api.d.ts') };
}

export function assertPublicApiEquivalent(local, published) {
  if (local?.jsSha256 !== published?.jsSha256 || local?.dtsSha256 !== published?.dtsSha256) {
    throw new Error('existing npm version has different public-api.js or public-api.d.ts; bump and publish npm');
  }
}

export function assertTarballIntegrity(bytes, integrity) {
  const actual = `sha512-${crypto.createHash('sha512').update(bytes).digest('base64')}`;
  if (actual !== integrity)
    throw new Error('downloaded npm tarball SHA-512 integrity does not match registry metadata');
  return actual;
}

function publicApiFromTarball(tarball) {
  const read = (name) =>
    sha256(execFileSync('tar', ['-xOf', tarball, `package/dist/npm/${name}`], { maxBuffer: 32 * 1024 * 1024 }));
  return { jsSha256: read('public-api.js'), dtsSha256: read('public-api.d.ts') };
}

export function createReleaseEnvelope({ releaseRoot, packageTarball, gitCommit, lifecycle = 'candidate' }) {
  if (!['candidate', 'supported', 'revoked'].includes(lifecycle)) {
    throw new Error(`Unsupported release lifecycle ${lifecycle}`);
  }
  const stable = readJson(path.join(releaseRoot, 'channels/stable-v5.json'), 'v5 release pointer').value;
  const releaseId = requireReleaseId(stable.releaseId);
  const sourceCommit = requireGitCommit(gitCommit, 'git commit');
  const manifestFile = path.join(releaseRoot, 'releases', releaseId, 'manifest.json');
  const manifest = readJson(manifestFile, 'v5 release manifest');
  if (manifest.value.version !== 5 || manifest.value.releaseId !== releaseId) {
    throw new Error('v5 release pointer and manifest do not agree');
  }
  const runtime = readJson(path.join(releaseRoot, '..', 'dist', 'onlyoffice-runtime-assets.json'), 'runtime manifest');
  const manifestSha256 = sha256(manifest.bytes);
  const runtimeManifestSha256 = sha256(runtime.bytes);
  if (
    stable.version !== 1 ||
    stable.manifestUrl !== `/releases/${releaseId}/manifest.json` ||
    stable.manifestSha256 !== manifestSha256
  ) {
    throw new Error('v5 release pointer does not bind the exact immutable manifest');
  }
  if (manifest.value.runtimeManifestSha256 !== runtimeManifestSha256) {
    throw new Error('v5 release manifest does not bind the runtime manifest');
  }
  if (manifest.value.sourceCommit !== sourceCommit) {
    throw new Error('v5 release manifest sourceCommit does not bind the candidate Git commit');
  }
  const envelope = {
    version: 1,
    lifecycle,
    runtime: {
      releaseId,
      manifestUrl: `/releases/${releaseId}/manifest.json`,
      manifestSha256: requireDigest(manifestSha256, 'release manifest digest'),
      runtimeManifestSha256: requireDigest(runtimeManifestSha256, 'runtime manifest digest'),
      hostBuildId: requireDigest(manifest.value.hostBuildId, 'Host build ID'),
      protocolHostBuildId: requireProtocolHostBuildId(
        manifest.value.protocolHostBuildId,
        manifest.value.packageVersion,
      ),
      sourceCommit,
    },
    npm: {
      packageName: '@agentbridges-ai/onlyoffice-browser',
      version: manifest.value.packageVersion,
      publicApi: publicApiIdentity(path.join(releaseRoot, '..', 'dist')),
      registry: { state: 'unverified' },
    },
    ...(packageTarball
      ? {
          sourcePackage: {
            tarball: path.basename(packageTarball),
            integrity: `sha512-${crypto.createHash('sha512').update(fs.readFileSync(packageTarball)).digest('base64')}`,
          },
        }
      : {}),
    x2t: manifest.value.x2t,
    source: {
      gitCommit: sourceCommit,
      lockfileSha256: requireDigest(sha256(fs.readFileSync('pnpm-lock.yaml')), 'pnpm lockfile digest'),
    },
  };
  return envelope;
}

// The candidate workflow invokes this before retaining its artifact.  It does
// not compare the entire npm tarball: package metadata can legitimately vary
// while the proxy API must not change under an already published version.
export async function bindExistingNpmPublication(envelope, { fetchImpl = fetch } = {}) {
  const registryUrl = `https://registry.npmjs.org/${encodeURIComponent(envelope.npm.packageName)}/${encodeURIComponent(envelope.npm.version)}`;
  const response = await fetchImpl(registryUrl, {
    cache: 'no-store',
    redirect: 'manual',
    signal: AbortSignal.timeout(15_000),
  });
  if (response.status === 404) return { ...envelope, npm: { ...envelope.npm, registry: { state: 'not-published' } } };
  if (!response.ok) throw new Error(`npm registry returned HTTP ${response.status}`);
  const metadata = await response.json();
  const identity = assertPublishedIdentity(metadata, {
    version: envelope.npm.version,
    gitHead: metadata.gitHead,
    integrity: metadata.dist?.integrity,
  });
  const attestationResponse = await fetchImpl(identity.attestationUrl, {
    cache: 'no-store',
    redirect: 'manual',
    signal: AbortSignal.timeout(15_000),
  });
  if (!attestationResponse.ok) throw new Error(`npm provenance returned HTTP ${attestationResponse.status}`);
  const attestation = await attestationResponse.json();
  assertNpmSlsaAttestation(attestation, {
    packageName: envelope.npm.packageName,
    version: envelope.npm.version,
    gitHead: metadata.gitHead,
    integrity: metadata.dist.integrity,
    releaseTag: `v${envelope.npm.version}`,
    workflow: '.github/workflows/release-npm.yml',
  });
  const tarballResponse = await fetchImpl(metadata.dist.tarball, {
    redirect: 'manual',
    signal: AbortSignal.timeout(30_000),
  });
  if (!tarballResponse.ok) throw new Error(`npm tarball returned HTTP ${tarballResponse.status}`);
  const temporary = path.join(os.tmpdir(), `onlyoffice-npm-${process.pid}-${crypto.randomUUID()}.tgz`);
  try {
    const tarballBytes = Buffer.from(await tarballResponse.arrayBuffer());
    assertTarballIntegrity(tarballBytes, metadata.dist.integrity);
    fs.writeFileSync(temporary, tarballBytes);
    const publishedApi = publicApiFromTarball(temporary);
    assertPublicApiEquivalent(envelope.npm.publicApi, publishedApi);
    return {
      ...envelope,
      npm: {
        ...envelope.npm,
        registry: {
          state: 'published',
          integrity: metadata.dist.integrity,
          gitHead: metadata.gitHead,
          attestationUrl: identity.attestationUrl,
          publicApi: publishedApi,
        },
      },
    };
  } finally {
    fs.rmSync(temporary, { force: true });
  }
}

export function validateReleaseEnvelope(envelope, { requireRegistry = true } = {}) {
  if (envelope?.version !== 1) throw new Error('release envelope version must be 1');
  if (!['candidate', 'supported', 'revoked'].includes(envelope.lifecycle))
    throw new Error('release envelope lifecycle is invalid');
  const releaseId = requireReleaseId(envelope.runtime?.releaseId);
  if (envelope.runtime?.manifestUrl !== `/releases/${releaseId}/manifest.json`) {
    throw new Error('release envelope manifest URL is invalid');
  }
  requireDigest(envelope.runtime?.manifestSha256, 'release manifest digest');
  requireDigest(envelope.runtime?.runtimeManifestSha256, 'runtime manifest digest');
  requireDigest(envelope.runtime?.hostBuildId, 'Host build ID');
  requireProtocolHostBuildId(envelope.runtime?.protocolHostBuildId, envelope.npm?.version);
  requireGitCommit(envelope.source?.gitCommit, 'git commit');
  requireGitCommit(envelope.runtime?.sourceCommit, 'runtime source commit');
  if (envelope.runtime.sourceCommit !== envelope.source.gitCommit) {
    throw new Error('release envelope runtime source commit does not match candidate source commit');
  }
  requireDigest(envelope.source?.lockfileSha256, 'pnpm lockfile digest');
  if (
    envelope.npm?.packageName !== '@agentbridges-ai/onlyoffice-browser' ||
    !/^\d+\.\d+\.\d+(?:[-+][0-9A-Za-z.-]+)?$/.test(envelope.npm?.version || '')
  ) {
    throw new Error('release envelope npm identity is invalid');
  }
  requireDigest(envelope.npm?.publicApi?.jsSha256, 'public-api.js digest');
  requireDigest(envelope.npm?.publicApi?.dtsSha256, 'public-api.d.ts digest');
  if (envelope.npm.registry?.state === 'published') {
    if (
      !/^sha512-[A-Za-z0-9+/]+={0,2}$/.test(envelope.npm.registry.integrity || '') ||
      !/^[a-f0-9]{40}$/.test(envelope.npm.registry.gitHead || '') ||
      typeof envelope.npm.registry.attestationUrl !== 'string'
    )
      throw new Error('release envelope npm registry identity is invalid');
    assertPublicApiEquivalent(envelope.npm.publicApi, envelope.npm.registry.publicApi);
  } else if (envelope.npm.registry?.state === 'not-published') {
    // A new version has no public npm identity yet; promotion can continue
    // because the current candidate is not reusing an immutable publication.
  } else if (!requireRegistry && envelope.npm.registry?.state === 'unverified') {
    // Synchronous unit callers may create an envelope before the CLI's network
    // enrichment step. Production promotion never accepts this state.
  } else {
    throw new Error('release envelope npm registry state is invalid');
  }
  if (
    envelope.sourcePackage &&
    (!/^[a-zA-Z0-9._+-]+\.tgz$/.test(envelope.sourcePackage.tarball || '') ||
      !/^sha512-[A-Za-z0-9+/]+={0,2}$/.test(envelope.sourcePackage.integrity || ''))
  ) {
    throw new Error('release envelope source package identity is invalid');
  }
  if (
    typeof envelope.x2t?.version !== 'string' ||
    !envelope.x2t.version ||
    !/^[a-f0-9]{40}$/.test(envelope.x2t?.commit || '') ||
    !/^[a-f0-9]{64}$/.test(envelope.x2t?.sha256 || '')
  ) {
    throw new Error('release envelope x2t identity is invalid');
  }
  return envelope;
}

function option(name) {
  const index = process.argv.indexOf(name);
  return index < 0 ? undefined : process.argv[index + 1];
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const output = option('--output');
  if (!output) throw new Error('--output is required');
  let envelope = createReleaseEnvelope({
    releaseRoot: path.resolve(option('--release-root') || '.onlyoffice-release'),
    packageTarball: option('--package-tarball'),
    gitCommit: option('--git-commit') || execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(),
    lifecycle: option('--lifecycle') || 'candidate',
  });
  envelope = await bindExistingNpmPublication(envelope);
  validateReleaseEnvelope(envelope);
  fs.mkdirSync(path.dirname(path.resolve(output)), { recursive: true });
  fs.writeFileSync(path.resolve(output), `${JSON.stringify(envelope, null, 2)}\n`);
}
