#!/usr/bin/env node

import crypto from 'node:crypto';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export const X2T_REPOSITORY = 'agentbridges-ai/onlyoffice-x2t-wasm';
export const X2T_RELEASE_WORKFLOW = '.github/workflows/build-release.yml';

const MAX_COMMAND_BUFFER = 256 * 1024 * 1024;

function sha(bytes, algorithm = 'sha256') {
  return crypto.createHash(algorithm).update(bytes).digest('hex');
}

function requireText(value, label) {
  if (typeof value !== 'string' || value.length === 0) throw new Error(`${label} is missing`);
  return value;
}

function requireCommit(value, label = 'x2t source commit') {
  if (!/^[a-f0-9]{40}$/.test(value || '')) throw new Error(`${label} must be a 40-character Git commit`);
  return value;
}

function normalizeTarEntry(value) {
  return value.replace(/^\.\//, '').replace(/\/$/, '');
}

export function expectedX2tAssets(version) {
  if (!/^\d+\.\d+\.\d+\+\d+$/.test(version || '')) throw new Error('x2t version is invalid');
  const archive = `onlyoffice-x2t-wasm-v${version}.tar.gz`;
  return {
    archive,
    archiveSha256: `${archive}.sha256`,
    archiveSha512: `${archive}.sha512`,
    filesSha256: 'x2t-files.sha256',
    filesSha512: 'x2t-files.sha512',
  };
}

export function verifyRepositoryMetadata(repository) {
  if (
    repository?.full_name !== X2T_REPOSITORY ||
    repository?.private !== false ||
    repository?.visibility !== 'public' ||
    repository?.archived === true
  ) {
    throw new Error('x2t source repository is not the expected active public repository');
  }
  return repository;
}

export function verifyTagIdentity(tagRef, tagObject, { tag, commit }) {
  requireCommit(commit);
  if (
    tagRef?.ref !== `refs/tags/${tag}` ||
    tagRef?.object?.type !== 'tag' ||
    !/^[a-f0-9]{40}$/.test(tagRef?.object?.sha || '')
  ) {
    throw new Error('x2t release tag is not an annotated tag');
  }
  if (
    tagObject?.tag !== tag ||
    tagObject?.object?.type !== 'commit' ||
    tagObject?.object?.sha !== commit ||
    tagObject?.verification?.verified !== true ||
    tagObject?.verification?.reason !== 'valid'
  ) {
    throw new Error('x2t release tag is not a verified signed tag for the expected commit');
  }
  return tagRef.object.sha;
}

export function verifyReleaseMetadata(release, { version, tag }) {
  const names = Object.values(expectedX2tAssets(version));
  if (
    release?.tag_name !== tag ||
    release?.draft !== false ||
    release?.prerelease !== false ||
    release?.immutable !== true
  ) {
    throw new Error('x2t release must be the exact immutable, non-draft, non-prerelease publication');
  }
  if (!Array.isArray(release.assets) || release.assets.length !== names.length) {
    throw new Error(`x2t release must contain exactly ${names.length} governed assets`);
  }
  const assets = new Map();
  for (const asset of release.assets) {
    if (
      !names.includes(asset?.name) ||
      assets.has(asset.name) ||
      !Number.isSafeInteger(asset?.id) ||
      asset.id < 1 ||
      !Number.isSafeInteger(asset?.size) ||
      asset.size < 1 ||
      asset?.state !== 'uploaded' ||
      !/^sha256:[a-f0-9]{64}$/.test(asset?.digest || '')
    ) {
      throw new Error('x2t release has a missing, duplicate, unexpected, or unverified asset');
    }
    assets.set(asset.name, asset);
  }
  if (names.some((name) => !assets.has(name))) throw new Error('x2t release asset set is incomplete');
  return assets;
}

export function verifyDownloadedAssets(directory, releaseAssets) {
  const expectedNames = [...releaseAssets.keys()].sort();
  const actualNames = fs.readdirSync(directory).sort();
  if (JSON.stringify(actualNames) !== JSON.stringify(expectedNames)) {
    throw new Error('downloaded x2t release asset set does not match GitHub metadata');
  }
  const bytesByName = new Map();
  for (const name of expectedNames) {
    const asset = releaseAssets.get(name);
    const file = path.join(directory, name);
    const stat = fs.lstatSync(file);
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size !== asset.size) {
      throw new Error(`downloaded x2t asset size or type mismatch: ${name}`);
    }
    const bytes = fs.readFileSync(file);
    if (`sha256:${sha(bytes)}` !== asset.digest) {
      throw new Error(`downloaded x2t asset GitHub digest mismatch: ${name}`);
    }
    bytesByName.set(name, bytes);
  }
  return bytesByName;
}

export function parseChecksumSidecar(bytes, algorithm) {
  const width = algorithm === 'sha256' ? 64 : algorithm === 'sha512' ? 128 : 0;
  if (!width) throw new Error(`unsupported checksum algorithm ${algorithm}`);
  const lines = bytes.toString('utf8').trimEnd().split('\n');
  const checksums = new Map();
  for (const line of lines) {
    const match = new RegExp(`^([a-f0-9]{${width}})  ([^\\s]+)$`).exec(line);
    if (!match || checksums.has(match[2])) throw new Error(`invalid or duplicate ${algorithm} sidecar entry`);
    checksums.set(match[2], match[1]);
  }
  if (checksums.size === 0) throw new Error(`${algorithm} sidecar is empty`);
  return checksums;
}

function requireExactChecksumSet(checksums, expectedNames, label) {
  if (
    checksums.size !== expectedNames.length ||
    expectedNames.some((name) => !checksums.has(name)) ||
    [...checksums.keys()].some((name) => !expectedNames.includes(name))
  ) {
    throw new Error(`${label} does not contain the exact expected files`);
  }
}

export function verifyChecksumSidecars(bytesByName, version) {
  const names = expectedX2tAssets(version);
  const archiveSha256 = parseChecksumSidecar(bytesByName.get(names.archiveSha256), 'sha256');
  const archiveSha512 = parseChecksumSidecar(bytesByName.get(names.archiveSha512), 'sha512');
  const filesSha256 = parseChecksumSidecar(bytesByName.get(names.filesSha256), 'sha256');
  const filesSha512 = parseChecksumSidecar(bytesByName.get(names.filesSha512), 'sha512');
  requireExactChecksumSet(archiveSha256, [names.archive], names.archiveSha256);
  requireExactChecksumSet(archiveSha512, [names.archive], names.archiveSha512);
  requireExactChecksumSet(filesSha256, ['build/x2t.js', 'build/x2t.wasm'], names.filesSha256);
  requireExactChecksumSet(filesSha512, ['build/x2t.js', 'build/x2t.wasm'], names.filesSha512);
  const archiveBytes = bytesByName.get(names.archive);
  if (archiveSha256.get(names.archive) !== sha(archiveBytes, 'sha256')) {
    throw new Error('x2t archive SHA-256 sidecar mismatch');
  }
  if (archiveSha512.get(names.archive) !== sha(archiveBytes, 'sha512')) {
    throw new Error('x2t archive SHA-512 sidecar mismatch');
  }
  return { names, archiveSha256, archiveSha512, filesSha256, filesSha512 };
}

function parseReleaseIdentity(bytes) {
  const values = new Map();
  for (const line of bytes.toString('utf8').trimEnd().split('\n')) {
    const match = /^([A-Za-z0-9_]+)=([^\s]+)$/.exec(line);
    if (!match || values.has(match[1])) throw new Error('x2t archive RELEASE metadata is malformed');
    values.set(match[1], match[2]);
  }
  return values;
}

export function verifyArchivePayload({
  entries,
  sourceCommit,
  release,
  x2tJs,
  x2tWasm,
  localRoot,
  checksums,
  tag,
  commit,
}) {
  const normalizedEntries = entries.map(normalizeTarEntry).filter(Boolean);
  for (const name of ['SOURCE_COMMIT', 'RELEASE', 'x2t.js', 'x2t.wasm']) {
    if (normalizedEntries.filter((entry) => entry === name).length !== 1) {
      throw new Error(`x2t archive must contain exactly one ${name}`);
    }
  }
  if (normalizedEntries.some((entry) => entry === '..' || entry.startsWith('../') || entry.includes('/../'))) {
    throw new Error('x2t archive contains an unsafe path');
  }
  if (sourceCommit.toString('utf8').trim() !== commit) throw new Error('x2t archive SOURCE_COMMIT mismatch');
  if (parseReleaseIdentity(release).get('version') !== tag) throw new Error('x2t archive RELEASE version mismatch');

  for (const [name, archiveBytes] of [
    ['x2t.js', x2tJs],
    ['x2t.wasm', x2tWasm],
  ]) {
    const localFile = path.join(localRoot, name);
    const stat = fs.lstatSync(localFile);
    if (!stat.isFile() || stat.isSymbolicLink()) throw new Error(`local public x2t asset is invalid: ${name}`);
    const localBytes = fs.readFileSync(localFile);
    if (!localBytes.equals(archiveBytes)) throw new Error(`local public ${name} differs from the attested archive`);
    if (
      checksums.filesSha256.get(`build/${name}`) !== sha(archiveBytes, 'sha256') ||
      checksums.filesSha512.get(`build/${name}`) !== sha(archiveBytes, 'sha512')
    ) {
      throw new Error(`x2t archive ${name} does not match its SHA-256/SHA-512 sidecars`);
    }
  }
  return {
    archiveSha256: checksums.archiveSha256.get(checksums.names.archive),
    x2tJsSha256: checksums.filesSha256.get('build/x2t.js'),
    x2tWasmSha256: checksums.filesSha256.get('build/x2t.wasm'),
  };
}

export function attestationVerifyArgs(archive, { tag, commit }) {
  return [
    'attestation',
    'verify',
    archive,
    '--repo',
    X2T_REPOSITORY,
    '--signer-workflow',
    `${X2T_REPOSITORY}/${X2T_RELEASE_WORKFLOW}`,
    '--signer-digest',
    commit,
    '--source-ref',
    `refs/tags/${tag}`,
    '--source-digest',
    commit,
    '--deny-self-hosted-runners',
    '--format',
    'json',
  ];
}

function sameSubjects(actual, expected) {
  if (!Array.isArray(actual) || actual.length !== expected.size) return false;
  const seen = new Set();
  for (const subject of actual) {
    if (seen.has(subject?.name) || subject?.digest?.sha256 !== expected.get(subject?.name)) return false;
    seen.add(subject.name);
  }
  return seen.size === expected.size;
}

export function verifyAttestationOutput(output, { tag, commit, archiveName, digests }) {
  let records;
  try {
    records = JSON.parse(output);
  } catch (error) {
    throw new Error('gh attestation verify did not return JSON', { cause: error });
  }
  const ref = `refs/tags/${tag}`;
  const repositoryUrl = `https://github.com/${X2T_REPOSITORY}`;
  const signer = `${repositoryUrl}/${X2T_RELEASE_WORKFLOW}@${ref}`;
  const expectedSubjects = new Map([
    [archiveName, digests.archiveSha256],
    ['x2t.js', digests.x2tJsSha256],
    ['x2t.wasm', digests.x2tWasmSha256],
  ]);
  const matched = Array.isArray(records)
    ? records.some(({ verificationResult }) => {
        const certificate = verificationResult?.signature?.certificate;
        const statement = verificationResult?.statement;
        const workflow = statement?.predicate?.buildDefinition?.externalParameters?.workflow;
        const dependencies = statement?.predicate?.buildDefinition?.resolvedDependencies;
        return (
          certificate?.subjectAlternativeName === signer &&
          certificate?.githubWorkflowRepository === X2T_REPOSITORY &&
          certificate?.githubWorkflowRef === ref &&
          certificate?.githubWorkflowSHA === commit &&
          certificate?.runnerEnvironment === 'github-hosted' &&
          certificate?.sourceRepositoryURI === repositoryUrl &&
          certificate?.sourceRepositoryDigest === commit &&
          certificate?.sourceRepositoryRef === ref &&
          certificate?.sourceRepositoryVisibilityAtSigning === 'public' &&
          statement?.predicateType === 'https://slsa.dev/provenance/v1' &&
          sameSubjects(statement?.subject, expectedSubjects) &&
          workflow?.repository === repositoryUrl &&
          workflow?.path === X2T_RELEASE_WORKFLOW &&
          workflow?.ref === ref &&
          Array.isArray(dependencies) &&
          dependencies.some(
            (dependency) =>
              dependency?.uri === `git+${repositoryUrl}@${ref}` && dependency?.digest?.gitCommit === commit,
          )
        );
      })
    : false;
  if (!matched)
    throw new Error('verified x2t attestation does not bind the exact archive, files, workflow, tag, and commit');
  return true;
}

function run(command, args, options = {}) {
  return execFileSync(command, args, { maxBuffer: MAX_COMMAND_BUFFER, ...options });
}

function option(argv, name) {
  const index = argv.indexOf(name);
  return index < 0 ? '' : argv[index + 1] || '';
}

export function verifyX2tPublication({ version, tag, commit, localRoot }, runCommand = run) {
  requireText(version, 'x2t version');
  requireText(tag, 'x2t release tag');
  requireCommit(commit);
  requireText(localRoot, 'local public x2t directory');
  if (tag !== `v${version}`) throw new Error('x2t version and release tag do not agree');
  const resolvedLocalRoot = path.resolve(localRoot);
  if (!fs.statSync(resolvedLocalRoot).isDirectory()) throw new Error('local public x2t directory is missing');

  const commandJson = (args) => JSON.parse(runCommand('gh', args, { encoding: 'utf8' }));
  verifyRepositoryMetadata(commandJson(['api', `repos/${X2T_REPOSITORY}`]));
  const refEndpoint = `repos/${X2T_REPOSITORY}/git/ref/tags/${encodeURIComponent(tag)}`;
  const tagRef = commandJson(['api', refEndpoint]);
  const tagObject = commandJson(['api', `repos/${X2T_REPOSITORY}/git/tags/${tagRef?.object?.sha || 'invalid'}`]);
  const tagObjectSha = verifyTagIdentity(tagRef, tagObject, { tag, commit });
  const release = commandJson(['api', `repos/${X2T_REPOSITORY}/releases/tags/${encodeURIComponent(tag)}`]);
  const releaseAssets = verifyReleaseMetadata(release, { version, tag });

  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-x2t-publication-'));
  try {
    runCommand('gh', ['release', 'download', tag, '--repo', X2T_REPOSITORY, '--dir', temporary]);
    const downloaded = verifyDownloadedAssets(temporary, releaseAssets);
    const checksums = verifyChecksumSidecars(downloaded, version);
    const archive = path.join(temporary, checksums.names.archive);
    const entries = runCommand('tar', ['-tzf', archive], { encoding: 'utf8' }).trimEnd().split('\n');
    const tarEntry = (name) => Buffer.from(runCommand('tar', ['-xOf', archive, `./${name}`]));
    const digests = verifyArchivePayload({
      entries,
      sourceCommit: tarEntry('SOURCE_COMMIT'),
      release: tarEntry('RELEASE'),
      x2tJs: tarEntry('x2t.js'),
      x2tWasm: tarEntry('x2t.wasm'),
      localRoot: resolvedLocalRoot,
      checksums,
      tag,
      commit,
    });
    const attestation = runCommand('gh', attestationVerifyArgs(archive, { tag, commit }), { encoding: 'utf8' });
    verifyAttestationOutput(attestation, { tag, commit, archiveName: checksums.names.archive, digests });

    const finalTagRef = commandJson(['api', refEndpoint]);
    if (finalTagRef?.object?.type !== 'tag' || finalTagRef?.object?.sha !== tagObjectSha) {
      throw new Error('x2t release tag changed during verification');
    }
    return { version, tag, commit, ...digests };
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const result = verifyX2tPublication({
    version: option(process.argv.slice(2), '--version'),
    tag: option(process.argv.slice(2), '--tag'),
    commit: option(process.argv.slice(2), '--commit'),
    localRoot: option(process.argv.slice(2), '--asset-root'),
  });
  process.stdout.write(
    `Verified immutable x2t publication ${result.tag} at ${result.commit} (${result.archiveSha256})\n`,
  );
}
