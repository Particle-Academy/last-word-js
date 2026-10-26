import { expect, it } from 'vitest';
import { deflateRawSync } from 'node:zlib';
import { crc32 } from '../src/zip/crc32';
import { Agent, ArchiveRefused, unzipSync } from '../src/index';
import { OdtReader } from '../src/reader/odt-reader';
import { DocxTemplate } from '../src/writer/docx-template';
import { mkdtempSync, openSync, ftruncateSync, closeSync, unlinkSync, rmdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const enc = new TextEncoder();
type Entry = { name: string; data: Uint8Array; method?: number; flags?: number; size?: number; crc?: number; descriptor?: boolean };
const text = (name: string, value: string, options: Partial<Entry> = {}): Entry => ({ name, data: enc.encode(value), ...options });
function fixture(entries: Entry[], comment = new Uint8Array()): Uint8Array {
  const locals: Uint8Array[] = [], central: Uint8Array[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = enc.encode(entry.name), method = entry.method ?? 0;
    const flags = entry.flags ?? (entry.descriptor ? 0x0808 : 0x0800);
    const data = method === 8 ? new Uint8Array(deflateRawSync(entry.data)) : entry.data;
    const crc = entry.crc ?? crc32(entry.data), size = entry.size ?? entry.data.length;
    const local = new Uint8Array(30 + name.length + data.length + (entry.descriptor ? 16 : 0));
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, flags, true); lv.setUint16(8, method, true);
    lv.setUint16(12, 33, true); lv.setUint32(14, entry.descriptor ? 0 : crc, true);
    lv.setUint32(18, entry.descriptor ? 0 : data.length, true); lv.setUint32(22, entry.descriptor ? 0 : size, true);
    lv.setUint16(26, name.length, true); local.set(name, 30); local.set(data, 30 + name.length);
    if (entry.descriptor) {
      const at = 30 + name.length + data.length;
      lv.setUint32(at, 0x08074b50, true); lv.setUint32(at + 4, crc, true);
      lv.setUint32(at + 8, data.length, true); lv.setUint32(at + 12, size, true);
    }
    const cd = new Uint8Array(46 + name.length), cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint16(8, flags, true); cv.setUint16(10, method, true); cv.setUint16(14, 33, true);
    cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true); cv.setUint32(24, size, true);
    cv.setUint16(28, name.length, true); cv.setUint32(42, offset, true); cd.set(name, 46);
    locals.push(local); central.push(cd); offset += local.length;
  }
  const cdSize = central.reduce((size, part) => size + part.length, 0);
  const out = new Uint8Array(offset + cdSize + 22 + comment.length);
  let pos = 0;
  for (const part of [...locals, ...central]) { out.set(part, pos); pos += part.length; }
  const end = new DataView(out.buffer, pos);
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true); end.setUint16(20, comment.length, true);
  out.set(comment, pos + 22); return out;
}

function readAttack(attack: string, odt = false): Uint8Array {
  const entries = odt ? [text('mimetype', 'application/vnd.oasis.opendocument.text'), text('content.xml', '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>')] : [
    text("word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>ok</w:t></w:r></w:p></w:body></w:document>"),
    text("word/styles.xml", "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>"),
  ];
  entries.push(attack === 'bomb' ? text('unknown/bomb.bin', 'a'.repeat(2097152), { method: 8 })
    : attack === 'duplicate' ? entries[0]! : attack === 'crc' ? text('unknown/corrupt.bin', 'opaque', { crc: 123 })
      : attack === 'path' ? text('../outside.xml', 'opaque') : text('unknown/valid.bin', 'opaque'));
  const bytes = fixture(entries); return attack === 'truncated' ? bytes.slice(0, -4) : bytes;
}
function refused(action: () => unknown, reason: string): void {
  let caught: unknown;
  try { action(); } catch (error) { caught = error; }
  expect(caught).toBeInstanceOf(ArchiveRefused);
  expect(caught).toMatchObject({ name: 'ArchiveRefused', reason });
  expect((caught as Error).constructor.name).toBe('ArchiveRefused');
}
it('accepts the valid package control', () => { expect(Agent.read(readAttack('valid'))).toHaveProperty('blocks'); });
for (const [attack, reason] of [["bomb","archive_limit"],["duplicate","archive_duplicate"],["crc","archive_corrupt"],["truncated","archive_corrupt"],["path","archive_path"]]) {
  it('unzipSync refuses ' + attack, () => refused(() => unzipSync(readAttack(attack!)), reason!));
  it('Agent.read refuses ' + attack, () => refused(() => Agent.read(readAttack(attack!)), reason!));
  it('ODT read refuses ' + attack, () => refused(() => new OdtReader().read(readAttack(attack!, true)), reason!));
  it('template read keeps the typed archive cause for ' + attack, () => {
    let caught: unknown; try { DocxTemplate.open(readAttack(attack!)); } catch (error) { caught = error; }
    expect(caught).toMatchObject({ name: 'TemplateException', reason: { name: 'ArchiveRefused', reason } });
  });
}
it('does not interpret an entry as an object prototype', () => { const parts = unzipSync(fixture([text('__proto__', 'opaque')])); expect(Object.getPrototypeOf(parts)).toBeNull(); expect(parts['__proto__']).toEqual(enc.encode('opaque')); });
it('bounds a Node template path before allocating its contents', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'sibling-template-')), input = join(dir, 'large.dotx');
  const fd = openSync(input, 'w'); ftruncateSync(fd, 67108865); closeSync(fd);
  try {
    let caught: unknown;
    try { await Agent.write({ blocks: [] }, join(dir, 'output.docx'), { template: input }); } catch (error) { caught = error; }
    expect(caught).toMatchObject({ name: 'TemplateException', reason: { name: 'ArchiveRefused', reason: 'archive_limit' } });
  } finally { unlinkSync(input); rmdirSync(dir); }
});
