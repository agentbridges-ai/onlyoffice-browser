import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
const publicationIdentity: any = await import('../../scripts/' + 'npm-publication-identity.mjs');
const { assertDistTagDoesNotRegress, assertNpmSlsaAttestation, assertPublishedIdentity, npmDistTag, npmPurl } =
  publicationIdentity;

const identity = {
  packageName: '@agentbridges-ai/onlyoffice-browser',
  version: '1.2.3',
  integrity: `sha512-${Buffer.from('package bytes').toString('base64')}`,
  gitHead: 'a'.repeat(40),
  releaseTag: 'v1.2.3',
  workflow: '.github/workflows/release-npm.yml',
};
const releaseWorkflow = fs.readFileSync('.github/workflows/release-npm.yml', 'utf8');

function attestation(overrides: Record<string, unknown> = {}) {
  const statement = {
    _type: 'https://in-toto.io/Statement/v1',
    predicateType: 'https://slsa.dev/provenance/v1',
    subject: [
      {
        name: npmPurl(identity.packageName, identity.version),
        digest: { sha512: Buffer.from(identity.integrity.replace('sha512-', ''), 'base64').toString('hex') },
      },
    ],
    predicate: {
      buildDefinition: {
        externalParameters: {
          workflow: {
            repository: 'https://github.com/agentbridges-ai/onlyoffice-browser',
            path: identity.workflow,
            ref: `refs/tags/${identity.releaseTag}`,
          },
        },
        resolvedDependencies: [
          {
            uri: `git+https://github.com/agentbridges-ai/onlyoffice-browser@refs/tags/${identity.releaseTag}`,
            digest: { gitCommit: identity.gitHead },
          },
        ],
      },
    },
    ...overrides,
  };
  return {
    attestations: [
      {
        bundle: {
          dsseEnvelope: {
            payloadType: 'application/vnd.in-toto+json',
            payload: Buffer.from(JSON.stringify(statement)).toString('base64'),
          },
        },
      },
    ],
  };
}

describe('npm publication identity', () => {
  it('keeps publication fail-closed and explicitly requests provenance', () => {
    expect(releaseWorkflow).toContain('response.status === 404');
    expect(releaseWorkflow).toContain('--provenance');
    expect(releaseWorkflow).toContain('npm pack --ignore-scripts');
    expect(releaseWorkflow).toContain('npm publish . --ignore-scripts');
    expect(releaseWorkflow).toContain('test "${GITHUB_REF}" = "refs/tags/${RELEASE_TAG}"');
    expect(releaseWorkflow).toContain('test "${GITHUB_SHA}" = "${RELEASE_COMMIT}"');
    expect(releaseWorkflow).toContain('REMOTE_TAG_OBJECT=');
    expect(releaseWorkflow).toContain("--jq '.verification.verified'");
    expect(releaseWorkflow).not.toContain("|| echo '{}' > dist-tags.json");
    expect(releaseWorkflow).toContain('npm registry returned');
  });

  it('keeps stable and prerelease dist-tags separate', () => {
    expect(npmDistTag('1.2.3')).toBe('latest');
    expect(npmDistTag('1.2.3-rc.1')).toBe('next');
    expect(() => assertDistTagDoesNotRegress({ latest: '2.0.0' }, '1.9.9')).toThrow(/move npm latest backwards/);
    expect(assertDistTagDoesNotRegress({ next: '1.2.3-rc.1' }, '1.2.3-rc.2')).toBe('next');
  });

  it('uses the npm scoped purl form with its slash intact', () => {
    expect(npmPurl(identity.packageName, identity.version)).toBe('pkg:npm/%40agentbridges-ai/onlyoffice-browser@1.2.3');
  });

  it('rejects immutable registry identity mismatch', () => {
    expect(() =>
      assertPublishedIdentity(
        { version: '1.2.3', gitHead: 'a', dist: { integrity: 'x', attestations: { url: 'u' } } },
        { version: '1.2.3', gitHead: 'b', integrity: 'x' },
      ),
    ).toThrow(/differs/);
    expect(() =>
      assertPublishedIdentity(
        {
          version: identity.version,
          gitHead: identity.gitHead,
          dist: { integrity: identity.integrity, attestations: { url: 'https://attacker.example/provenance' } },
        },
        identity,
      ),
    ).toThrow(/not registry-owned/);
  });

  it('decodes npm DSSE and binds purl, sha512, commit, tag and workflow', () => {
    expect(assertNpmSlsaAttestation(attestation(), identity)).toMatchObject({
      predicateType: 'https://slsa.dev/provenance/v1',
    });
    expect(() => assertNpmSlsaAttestation(attestation(), { ...identity, integrity: 'not-sri' })).toThrow(/SHA-512 SRI/);
    expect(() => assertNpmSlsaAttestation(attestation({ subject: [] }), identity)).toThrow(/purl and sha512/);
    expect(() =>
      assertNpmSlsaAttestation(
        attestation({
          predicate: {
            buildDefinition: {
              externalParameters: {
                workflow: {
                  repository: 'https://github.com/attacker/repository',
                  path: identity.workflow,
                  ref: `refs/tags/${identity.releaseTag}`,
                },
              },
              resolvedDependencies: [
                {
                  uri: `git+https://github.com/agentbridges-ai/onlyoffice-browser@refs/tags/${identity.releaseTag}`,
                  digest: { gitCommit: identity.gitHead },
                },
              ],
            },
          },
        }),
        identity,
      ),
    ).toThrow(/exact release workflow/);
  });
});
