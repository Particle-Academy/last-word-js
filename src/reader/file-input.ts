import { ArchiveRefused } from "../exceptions";
import { DEFAULT_ARCHIVE_LIMITS } from "../zip/source-archive";

/** Node-only file input, bounded before allocation and even if the file grows. */
export async function readFileBounded(path: string): Promise<Uint8Array> {
  const fs = await import("node:fs");
  if (!fs.statSync(path).isFile()) throw new ArchiveRefused("archive_path");
  const fd = fs.openSync(path, "r");
  const limit = DEFAULT_ARCHIVE_LIMITS.maxArchiveBytes;
  try {
    const stat = fs.fstatSync(fd);
    if (!stat.isFile()) throw new ArchiveRefused("archive_path");
    if (stat.size > limit) throw new ArchiveRefused("archive_limit");
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const chunk = new Uint8Array(Math.min(65536, limit + 1 - total));
      const count = fs.readSync(fd, chunk, 0, chunk.length, null);
      if (count === 0) break;
      if (total + count > limit) throw new ArchiveRefused("archive_limit");
      chunks.push(chunk.subarray(0, count)); total += count;
    }
    const bytes = new Uint8Array(total);
    let at = 0;
    for (const chunk of chunks) { bytes.set(chunk, at); at += chunk.length; }
    return bytes;
  } finally { fs.closeSync(fd); }
}
