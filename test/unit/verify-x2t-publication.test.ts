import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';

const x2tPublication: any = await import('../../scripts/' + 'verify-x2t-publication.mjs');
const {
  X2T_RELEASE_WORKFLOW,
  X2T_REPOSITORY,
  attestationVerifyArgs,
  expectedX2tAssets,
  verifyArchivePayload,
  verifyAttestationOutput,
  verifyChecksumSidecars,
  verifyDownloadedAssets,
  verifyReleaseMetadata,
  verifyRepositoryMetadata,
  verifyTagIdentity,
} = x2tPublication;

const version = '9.3.0+2';
const tag = `v${version}`;
const commit = '1bb9b45a399f87ca162eea0c86abd4660f295469';
const temporaryDirectories: string[] = [];
const sha = (bytes: Uint8Array | string, algorithm = 'sha256') =>
  crypto.createHash(algorithm).update(bytes).digest('hex');

function temp() {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-x2t-verifier-'));
  temporaryDirectories.push(directory);
  return directory;
}

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) fs.rmSync(directory, { recursive: true, force: true });
});

function releaseFixture(root: string) {
  const names = expectedX2tAssets(version);
  const archive = Buffer.from('archive bytes');
  const x2tJs = Buffer.from('x2t javascript');
  const x2tWasm = Buffer.from('x2t wasm');
  const values = new Map<string, Buffer>([
    [names.archive, archive],
    [names.archiveSha256, Buffer.from(`${sha(archive)}  ${names.archive}\n`)],
    [names.archiveSha512, Buffer.from(`${sha(archive, 'sha512')}  ${names.archive}\n`)],
    [names.filesSha256, Buffer.from(`${sha(x2tJs)}  build/x2t.js\n${sha(x2tWasm)}  build/x2t.wasm\n`)],
    [
      names.filesSha512,
      Buffer.from(`${sha(x2tJs, 'sha512')}  build/x2t.js\n${sha(x2tWasm, 'sha512')}  build/x2t.wasm\n`),
    ],
  ]);
  const assets = [...values].map(([name, bytes], index) => {
    fs.writeFileSync(path.join(root, name), bytes);
    return { id: index + 1, name, size: bytes.byteLength, state: 'uploaded', digest: `sha256:${sha(bytes)}` };
  });
  return { names, archive, x2tJs, x2tWasm, assets };
}

function attestation(archiveName: string, digests: Record<string, string>, ref = `refs/tags/${tag}`) {
  const repositoryUrl = `https://github.com/${X2T_REPOSITORY}`;
  return JSON.stringify([
    {
      verificationResult: {
        signature: {
          certificate: {
            subjectAlternativeName: `${repositoryUrl}/${X2T_RELEASE_WORKFLOW}@${ref}`,
            githubWorkflowRepository: X2T_REPOSITORY,
            githubWorkflowRef: ref,
            githubWorkflowSHA: commit,
            runnerEnvironment: 'github-hosted',
            sourceRepositoryURI: repositoryUrl,
            sourceRepositoryDigest: commit,
            sourceRepositoryRef: ref,
            sourceRepositoryVisibilityAtSigning: 'public',
          },
        },
        statement: {
          predicateType: 'https://slsa.dev/provenance/v1',
          subject: [
            { name: archiveName, digest: { sha256: digests.archiveSha256 } },
            { name: 'x2t.js', digest: { sha256: digests.x2tJsSha256 } },
            { name: 'x2t.wasm', digest: { sha256: digests.x2tWasmSha256 } },
          ],
          predicate: {
            buildDefinition: {
              externalParameters: {
                workflow: { repository: repositoryUrl, path: X2T_RELEASE_WORKFLOW, ref },
              },
              resolvedDependencies: [{ uri: `git+${repositoryUrl}@${ref}`, digest: { gitCommit: commit } }],
            },
          },
        },
      },
    },
  ]);
}

describe('public x2t publication verifier', () => {
  it('requires the exact public repository, verified annotated tag, and immutable release asset set', () => {
    expect(
      verifyRepositoryMetadata({
        full_name: X2T_REPOSITORY,
        private: false,
        visibility: 'public',
        archived: false,
      }),
    ).toBeTruthy();
    expect(
      verifyTagIdentity(
        { ref: `refs/tags/${tag}`, object: { type: 'tag', sha: 'a'.repeat(40) } },
        {
          tag,
          object: { type: 'commit', sha: commit },
          verification: { verified: true, reason: 'valid' },
        },
        { tag, commit },
      ),
    ).toBe('a'.repeat(40));

    const root = temp();
    const { assets } = releaseFixture(root);
    expect(
      verifyReleaseMetadata(
        { tag_name: tag, draft: false, prerelease: false, immutable: true, assets },
        { version, tag },
      ).size,
    ).toBe(5);
    expect(() =>
      verifyReleaseMetadata(
        { tag_name: tag, draft: false, prerelease: false, immutable: false, assets },
        { version, tag },
      ),
    ).toThrow(/immutable/);
  });

  it('checks GitHub asset digests, both checksum families, archive identity, and repository x2t bytes', () => {
    const root = temp();
    const localRoot = temp();
    const { x2tJs, x2tWasm, assets } = releaseFixture(root);
    fs.writeFileSync(path.join(localRoot, 'x2t.js'), x2tJs);
    fs.writeFileSync(path.join(localRoot, 'x2t.wasm'), x2tWasm);
    const releaseAssets = verifyReleaseMetadata(
      { tag_name: tag, draft: false, prerelease: false, immutable: true, assets },
      { version, tag },
    );
    const downloaded = verifyDownloadedAssets(root, releaseAssets);
    const checksums = verifyChecksumSidecars(downloaded, version);
    const digests = verifyArchivePayload({
      entries: ['./', './RELEASE', './SOURCE_COMMIT', './x2t.js', './x2t.wasm'],
      sourceCommit: Buffer.from(`${commit}\n`),
      release: Buffer.from(`version=${tag}\nonlyoffice_core=v9.3.0.140\nemscripten=4.0.11\n`),
      x2tJs,
      x2tWasm,
      localRoot,
      checksums,
      tag,
      commit,
    });
    expect(digests).toEqual({
      archiveSha256: sha(Buffer.from('archive bytes')),
      x2tJsSha256: sha(x2tJs),
      x2tWasmSha256: sha(x2tWasm),
    });

    const sidecar = path.join(root, expectedX2tAssets(version).filesSha256);
    const tampered = fs.readFileSync(sidecar);
    tampered[0] = tampered[0] === 0x30 ? 0x31 : 0x30;
    fs.writeFileSync(sidecar, tampered);
    expect(() => verifyDownloadedAssets(root, releaseAssets)).toThrow(/GitHub digest mismatch/);
  });

  it('binds gh verification to the exact repository, workflow, tag, commit, and three subjects', () => {
    const names = expectedX2tAssets(version);
    const digests = {
      archiveSha256: 'a'.repeat(64),
      x2tJsSha256: 'b'.repeat(64),
      x2tWasmSha256: 'c'.repeat(64),
    };
    expect(attestationVerifyArgs('/tmp/archive.tgz', { tag, commit })).toEqual([
      'attestation',
      'verify',
      '/tmp/archive.tgz',
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
    ]);
    expect(
      verifyAttestationOutput(attestation(names.archive, digests), {
        tag,
        commit,
        archiveName: names.archive,
        digests,
      }),
    ).toBe(true);
    expect(() =>
      verifyAttestationOutput(attestation(names.archive, digests, 'refs/tags/v9.3.0+3'), {
        tag,
        commit,
        archiveName: names.archive,
        digests,
      }),
    ).toThrow(/does not bind/);
  });
});
