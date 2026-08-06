import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { deflateRawSync } from 'node:zlib';
import { afterEach, describe, expect, it } from 'vitest';
import {
  extractArtifactEntry,
  parseZipCentralDirectory,
  selectArtifactEntries,
  // @ts-expect-error JavaScript release helper intentionally has no declaration output.
} from '../../scripts/download-github-artifact-files.mjs';

const temporaryDirectories: string[] = [];

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function zipFixture(name: string, content: Buffer, compressed: boolean): Buffer {
  const payload = compressed ? deflateRawSync(content) : content;
  const local = Buffer.alloc(30 + Buffer.byteLength(name));
  local.writeUInt32LE(0x04034b50, 0);
  local.writeUInt16LE(20, 4);
  local.writeUInt16LE(compressed ? 8 : 0, 8);
  local.writeUInt32LE(crc32(content), 14);
  local.writeUInt32LE(payload.byteLength, 18);
  local.writeUInt32LE(content.byteLength, 22);
  local.writeUInt16LE(Buffer.byteLength(name), 26);
  local.write(name, 30);
  const central = Buffer.alloc(46 + Buffer.byteLength(name));
  central.writeUInt32LE(0x02014b50, 0);
  central.writeUInt16LE(20, 4);
  central.writeUInt16LE(20, 6);
  central.writeUInt16LE(compressed ? 8 : 0, 10);
  central.writeUInt32LE(crc32(content), 16);
  central.writeUInt32LE(payload.byteLength, 20);
  central.writeUInt32LE(content.byteLength, 24);
  central.writeUInt16LE(Buffer.byteLength(name), 28);
  central.write(name, 46);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(1, 8);
  eocd.writeUInt16LE(1, 10);
  eocd.writeUInt32LE(central.byteLength, 12);
  eocd.writeUInt32LE(local.byteLength + payload.byteLength, 16);
  return Buffer.concat([local, payload, central, eocd]);
}

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) fs.rmSync(directory, { recursive: true, force: true });
});

describe('GitHub artifact range reader', () => {
  it('parses and selects safe ZIP entries without broadening a glob into directories', () => {
    const bytes = zipFixture('dist/office-host.js', Buffer.from('host'), false);
    const centralOffset = bytes.lastIndexOf(Buffer.from('PK\x01\x02', 'binary'));
    const entries = parseZipCentralDirectory(bytes.subarray(centralOffset, bytes.length - 22));
    expect(entries[0]).toMatchObject({ name: 'dist/office-host.js', method: 0, uncompressedSize: 4 });
    expect(
      selectArtifactEntries([...entries, { ...entries[0], name: 'dist/', directory: true }], { includes: ['dist/**'] }),
    ).toHaveLength(1);
  });

  it('range-extracts and verifies a deflated entry before writing it', async () => {
    const content = Buffer.from('incremental artifact payload '.repeat(100));
    const archiveBytes = zipFixture('release-evidence/payload.txt', content, true);
    const centralOffset = archiveBytes.lastIndexOf(Buffer.from('PK\x01\x02', 'binary'));
    const [entry] = parseZipCentralDirectory(archiveBytes.subarray(centralOffset, archiveBytes.length - 22));
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-artifact-range-'));
    temporaryDirectories.push(directory);
    const fetchImpl = async (_url: string, init: { headers?: { range?: string } }) => {
      const match = /bytes=(\d+)-(\d+)/.exec(init.headers?.range || '');
      if (!match) throw new Error('range header missing');
      const start = Number(match[1]);
      const end = Number(match[2]);
      const payload = archiveBytes.subarray(start, end + 1);
      const exact = new ArrayBuffer(payload.byteLength);
      new Uint8Array(exact).set(payload);
      return new Response(exact, {
        status: 206,
        headers: { 'content-range': `bytes ${start}-${end}/${archiveBytes.byteLength}` },
      });
    };
    const result = await extractArtifactEntry({
      archive: { url: 'https://artifact.test/archive.zip', stats: { rangeBytes: 0 } },
      entry,
      outputRoot: directory,
      fetchImpl,
    });
    const output = fs.readFileSync(path.join(directory, entry.name));
    expect(output.equals(content)).toBe(true);
    expect(result).toMatchObject({ path: entry.name, bytes: content.byteLength });
    expect(result.sha256).toBe(crypto.createHash('sha256').update(content).digest('hex'));
  });
});
