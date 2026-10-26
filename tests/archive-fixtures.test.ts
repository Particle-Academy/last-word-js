import { expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { Agent, unzipSync } from '../src/index';
import { SourceArchive } from '../src/zip/source-archive';
it('retains every fixture payload and pre-patch reader projection', () => {
  const digest = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex');
  const cases = JSON.parse(readFileSync(resolve('tests/fixtures/archive-baseline.json'), 'utf8')) as Record<string, { sourceSha256: string; entries: number; payloadsSha256: string; projectionSha256: string }>;
  for (const [file, expected] of Object.entries(cases)) {
    const bytes = new Uint8Array(readFileSync(resolve(file))), files = unzipSync(bytes);
    const parts = Object.keys(files).sort().map(name => [name, digest(files[name]!)]);
    expect(digest(bytes)).toBe(expected.sourceSha256); expect(parts.length).toBe(expected.entries);
    expect(digest(JSON.stringify(parts))).toBe(expected.payloadsSha256);
    expect(digest(JSON.stringify(Agent.read(bytes)))).toBe(expected.projectionSha256);
    expect(SourceArchive.open(bytes).toBytes()).toEqual(bytes);
  }
});
