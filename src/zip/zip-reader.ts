import { SourceArchive, type ArchiveLimits } from "./source-archive";

/** STORE/DEFLATE reader with metadata, CRC, duplicate and expansion checks. */
export function unzipSync(data: Uint8Array, limits: ArchiveLimits = {}): Record<string, Uint8Array> {
  const source = SourceArchive.open(data, limits);
  const result: Record<string, Uint8Array> = Object.create(null);
  for (const name of source.names()) result[name] = source.read(name)!;
  return result;
}

/** Selected parts still pass whole-archive validation, including ignored entries. */
export function unzipEntries(data: Uint8Array, names: string[], maxBytes: number): Record<string, Uint8Array> {
  const source = SourceArchive.open(data, { maxPartBytes: maxBytes });
  const result: Record<string, Uint8Array> = Object.create(null);
  for (const name of names) {
    const bytes = source.read(name);
    if (bytes !== null) result[name] = bytes;
  }
  return result;
}
