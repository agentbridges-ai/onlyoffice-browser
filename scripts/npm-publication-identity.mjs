import crypto from 'node:crypto';
import fs from 'node:fs';

export const sha512Integrity = (file) =>
  `sha512-${crypto.createHash('sha512').update(fs.readFileSync(file)).digest('base64')}`;

function requireText(value, label) {
  if (typeof value !== 'string' || !value) throw new Error(`${label} is missing`);
  return value;
}

function versionParts(version) {
  const match = /^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+[0-9A-Za-z.-]+)?$/.exec(version);
  if (!match) throw new Error(`invalid SemVer ${version}`);
  return { numeric: match.slice(1, 4).map(Number), prerelease: match[4]?.split('.') || [] };
}

export const npmDistTag = (version) => (versionParts(version).prerelease.length ? 'next' : 'latest');

// This deliberately implements only ordering needed for npm dist-tag guards;
// it rejects non-SemVer tags instead of guessing how npm would order them.
export function compareSemver(left, right) {
  const a = versionParts(left);
  const b = versionParts(right);
  for (let i = 0; i < 3; i += 1) if (a.numeric[i] !== b.numeric[i]) return a.numeric[i] > b.numeric[i] ? 1 : -1;
  if (!a.prerelease.length || !b.prerelease.length) return a.prerelease.length ? -1 : b.prerelease.length ? 1 : 0;
  const length = Math.max(a.prerelease.length, b.prerelease.length);
  for (let i = 0; i < length; i += 1) {
    if (a.prerelease[i] === undefined) return -1;
    if (b.prerelease[i] === undefined) return 1;
    if (a.prerelease[i] === b.prerelease[i]) continue;
    const an = /^\d+$/.test(a.prerelease[i]);
    const bn = /^\d+$/.test(b.prerelease[i]);
    if (an && bn) return Number(a.prerelease[i]) > Number(b.prerelease[i]) ? 1 : -1;
    if (an !== bn) return an ? -1 : 1;
    return a.prerelease[i] > b.prerelease[i] ? 1 : -1;
  }
  return 0;
}

export function assertDistTagDoesNotRegress(distTags, version) {
  const tag = npmDistTag(version);
  const current = distTags?.[tag];
  if (current && compareSemver(version, current) < 0) {
    throw new Error(`refusing to move npm ${tag} backwards from ${current} to ${version}`);
  }
  return tag;
}

export function assertPublishedIdentity(metadata, { version, gitHead, integrity }) {
  if (metadata.version !== version || metadata.gitHead !== gitHead || metadata.dist?.integrity !== integrity)
    throw new Error('published npm identity differs from the signed release');
  const slsa = metadata.dist?.attestations?.url;
  if (!slsa) throw new Error('published npm provenance record is missing');
  let attestationUrl;
  try {
    attestationUrl = new URL(slsa);
  } catch (error) {
    throw new Error('published npm provenance URL is invalid', { cause: error });
  }
  if (
    attestationUrl.protocol !== 'https:' ||
    attestationUrl.hostname !== 'registry.npmjs.org' ||
    attestationUrl.username ||
    attestationUrl.password
  ) {
    throw new Error('published npm provenance URL is not registry-owned');
  }
  return { integrity: metadata.dist.integrity, attestationUrl: slsa };
}

function decodedStatements(document) {
  const records = Array.isArray(document?.attestations) ? document.attestations : [document];
  return records.flatMap((record) => {
    const envelope = record?.bundle?.dsseEnvelope || record?.dsseEnvelope || record?.envelope;
    if (!envelope?.payload || envelope.payloadType !== 'application/vnd.in-toto+json') return [];
    try {
      return [JSON.parse(Buffer.from(envelope.payload, 'base64').toString('utf8'))];
    } catch {
      return [];
    }
  });
}

export function npmPurl(packageName, version) {
  const match = /^@([^/]+)\/([^/]+)$/.exec(packageName);
  if (!match) throw new Error(`expected scoped npm package name, got ${packageName}`);
  // npm's SLSA subject encodes the @ but deliberately keeps the namespace slash.
  return `pkg:npm/%40${match[1]}/${match[2]}@${version}`;
}

export function assertNpmSlsaAttestation(
  document,
  {
    packageName,
    version,
    integrity,
    gitHead,
    releaseTag,
    workflow,
    repository = 'https://github.com/agentbridges-ai/onlyoffice-browser',
  },
) {
  const expectedPurl = npmPurl(packageName, version);
  const expectedIntegrity = requireText(integrity, 'package integrity');
  if (!/^sha512-[A-Za-z0-9+/]+={0,2}$/.test(expectedIntegrity)) {
    throw new Error('package integrity is not a SHA-512 SRI');
  }
  const expectedSha512 = expectedIntegrity.slice('sha512-'.length);
  const expectedSha512Hex = Buffer.from(expectedSha512, 'base64').toString('hex');
  const statement = decodedStatements(document).find((candidate) => {
    const subject = candidate?.subject?.find((item) => item?.name === expectedPurl);
    const digest = subject?.digest?.sha512;
    return (
      candidate?._type === 'https://in-toto.io/Statement/v1' &&
      candidate?.predicateType === 'https://slsa.dev/provenance/v1' &&
      digest === expectedSha512Hex
    );
  });
  if (!statement) throw new Error('npm SLSA DSSE does not bind the scoped purl and sha512 subject');
  const tagRef = `refs/tags/${releaseTag}`;
  const buildDefinition = statement.predicate?.buildDefinition;
  if (
    !buildDefinition?.resolvedDependencies?.some(
      (dependency) => dependency?.uri === `git+${repository}@${tagRef}` && dependency?.digest?.gitCommit === gitHead,
    )
  ) {
    throw new Error('npm SLSA DSSE does not bind source commit and signed release tag');
  }
  const workflowIdentity = buildDefinition.externalParameters?.workflow;
  if (
    workflowIdentity?.repository !== repository ||
    workflowIdentity?.path !== workflow ||
    workflowIdentity?.ref !== tagRef
  ) {
    throw new Error('npm SLSA DSSE does not bind the exact release workflow');
  }
  return statement;
}
