import { readFile, readdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import packageJson from '../package.json' with { type: 'json' };

const assetsRoot = resolve(import.meta.dirname, '../dist/assets');
const hostBundleName = (await readdir(assetsRoot)).find((name) => /^officeHost-.*\.js$/.test(name));
if (!hostBundleName) throw new Error('Built Office Host bundle is missing');

const hostBundle = await readFile(resolve(assetsRoot, hostBundleName), 'utf8');
const expectedBuildId = `office-host-${packageJson.version}-r1`;
const versionImport = /from["']\.\/(version-[^"']+\.js)["']/.exec(hostBundle)?.[1];
const versionBundle = versionImport
  ? await readFile(resolve(assetsRoot, versionImport), 'utf8')
  : '';
if (
  !versionImport ||
  !versionBundle.includes(packageJson.version) ||
  !hostBundle.includes('office-host-${')
) {
  throw new Error(`Built Office Host identity does not match npm ${packageJson.version}/${expectedBuildId}`);
}

console.log(`Verified Office Host identity: ${packageJson.version}/${expectedBuildId}`);
