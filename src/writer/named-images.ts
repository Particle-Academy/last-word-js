/**
 * Images the HOST registers with the writer, by name (CC-fancy-0039), so a model
 * places `{"type": "image", "src": "image:logo"}` instead of carrying the bytes --
 * or inventing an image -- and a letterhead's logo goes in a header part as easily
 * as in the body. Mirrors PHP `LastWord\Writer\NamedImages`: the same options, the
 * same checks in the same order, the same refusal codes, the same guide and schema
 * text. The shared fixture (test/fixtures/named-images.json) pins all of it.
 *
 * - `images`: `{name: image}`, an image being PNG or JPEG bytes (`Uint8Array` /
 *   `ArrayBuffer`), a PNG or JPEG data URL, or `{data, widthPx?, heightPx?, alt?}`
 *   with one of those as `data`. The size and alt are what a block naming the image
 *   gets unless it sets its own: a block setting either `widthPx` or `heightPx`
 *   takes neither registered size, so the image keeps its shape.
 * - `imageLimits`: `{maxBytes?, maxWidthPx?, maxHeightPx?, maxImages?}`, positive
 *   integers; each absent one is the default in `DEFAULT_IMAGE_LIMITS`.
 *
 * A name is resolved when the document is written. The writer never sees a name: it
 * is given the data URL, so a named image writes exactly what the same image as a
 * data URL writes.
 */

import { ImageRefused, SchemaException } from "../exceptions";
import { base64Encode, jpegSize } from "../helpers/image-size";
import type { ValidationError } from "../schema/types";
import { isPlainObject } from "../util";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** What a block's `src` starts with to name a registered image. */
export const IMAGE_PREFIX = "image:";

/** A letter, then up to 63 letters, digits, `-` or `_`. */
const NAME_PATTERN = /^[A-Za-z][A-Za-z0-9_-]{0,63}$/;

/** Secure defaults; a host raises or lowers each through `imageLimits`. */
export const DEFAULT_IMAGE_LIMITS = {
  maxBytes: 5242880,
  maxWidthPx: 4096,
  maxHeightPx: 4096,
  maxImages: 32,
} as const;

export type ImageLimits = { maxBytes?: number; maxWidthPx?: number; maxHeightPx?: number; maxImages?: number };

/** One image as a host registers it. */
export type NamedImage =
  | Uint8Array
  | ArrayBuffer
  | string
  | { data: Uint8Array | ArrayBuffer | string; widthPx?: number; heightPx?: number; alt?: string };

export type ImageOptions = { images?: Record<string, NamedImage> | null; imageLimits?: ImageLimits | null };

type Limits = { maxBytes: number; maxWidthPx: number; maxHeightPx: number; maxImages: number };

type Registered = {
  type: "png" | "jpeg";
  bytes: Uint8Array;
  width: number;
  height: number;
  widthPx?: number;
  heightPx?: number;
  alt?: string;
};

const IMAGE_KEYS = ["data", "widthPx", "heightPx", "alt"];

const has = (o: object, k: string): boolean => Object.prototype.hasOwnProperty.call(o, k);

const isBytes = (v: unknown): v is Uint8Array | ArrayBuffer => v instanceof Uint8Array || v instanceof ArrayBuffer;

const toU8 = (v: Uint8Array | ArrayBuffer): Uint8Array => (v instanceof Uint8Array ? v : new Uint8Array(v));

function startsWith(bytes: Uint8Array, prefix: number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((b, i) => bytes[i] === b);
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const JPEG_SIGNATURE = [0xff, 0xd8, 0xff];

/**
 * PNG IHDR width/height as PHP's `ImageSize::png()` reads them: unsigned, so a
 * header claiming 2^31 px or more is too large rather than unreadable, the same
 * answer the PHP engine gives.
 */
function pngDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  if (bytes.length < 24 || !startsWith(bytes, PNG_SIGNATURE)) return null;
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52) return null;
  const u32 = (at: number): number => ((bytes[at]! << 24) | (bytes[at + 1]! << 16) | (bytes[at + 2]! << 8) | bytes[at + 3]!) >>> 0;
  const width = u32(16);
  const height = u32(20);
  if (width < 1 || height < 1) return null;
  return { width, height };
}

/** Strict base64 decoding of a payload already checked against the alphabet. */
function decodeBase64(payload: string): Uint8Array {
  if (typeof Buffer !== "undefined") {
    const buf = Buffer.from(payload, "base64");
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  }
  // eslint-disable-next-line no-undef
  const bin = atob(payload);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export class NamedImages {
  private constructor(private readonly images: Map<string, Registered>) {}

  /**
   * The `images` / `imageLimits` options as registered images, or null when no image
   * is registered (absent, null or empty) -- which is then exactly the behaviour from
   * before this existed. Throws `ImageRefused`.
   */
  static fromOptions(options: ImageOptions | undefined | null): NamedImages | null {
    const limits = NamedImages.limits(options?.imageLimits);
    const images: unknown = options?.images;
    if (images === undefined || images === null) return null;
    // An empty JSON array is PHP's empty object too.
    if (Array.isArray(images) && images.length === 0) return null;
    if (typeof images !== "object" || isBytes(images)) {
      throw new ImageRefused("image_option_invalid", "images must be an object of name => image.");
    }
    const entries = Object.entries(images as Record<string, unknown>);
    if (entries.length === 0) return null;
    if (entries.length > limits.maxImages) {
      throw new ImageRefused("image_count_exceeded", `at most ${limits.maxImages} images may be registered (imageLimits.maxImages).`);
    }
    const registered = new Map<string, Registered>();
    for (const [name, image] of entries) {
      if (!NAME_PATTERN.test(name)) {
        throw new ImageRefused("image_name_invalid", "an image name is a letter followed by up to 63 letters, digits, - or _.");
      }
      registered.set(name, NamedImages.register(name, image, limits));
    }
    return new NamedImages(registered);
  }

  /** The registered names, in registration order. */
  names(): string[] {
    return [...this.images.keys()];
  }

  /**
   * The document with every `image:<name>` src replaced by its data URL, and the
   * registered size and alt filled in where the block sets none. Never modifies the
   * input. Throws `SchemaException` for a name that is not registered (validation
   * reports it first).
   */
  resolve<T>(doc: T): T {
    return this.walk(doc, "") as T;
  }

  /** The section `Agent.guide()` appends. Identical to PHP's. */
  guideSection(): string {
    let lines = "";
    for (const [name, image] of this.images) {
      let line = "- `" + IMAGE_PREFIX + name + "`: " + image.type.toUpperCase() + ", " + image.width + " x " + image.height + " px";
      if (image.widthPx !== undefined && image.heightPx !== undefined) {
        line += "; placed at " + image.widthPx + " x " + image.heightPx + " px";
      } else if (image.widthPx !== undefined) {
        line += "; placed " + image.widthPx + " px wide";
      } else if (image.heightPx !== undefined) {
        line += "; placed " + image.heightPx + " px high";
      }
      if (image.alt !== undefined) line += "; alt " + JSON.stringify(image.alt);
      lines += line + "\n";
    }
    return (
      "\n## Images you can place\n\n" +
      'The host registered these images for this document. Place one with `{"type": "image", "src": "image:NAME"}`, ' +
      "in the body or in a header or footer that holds `blocks` (a letterhead puts its logo in `page.header` or `page.firstHeader`). " +
      "Use only these names: never invent an image, and never write out a data URL for one of them. " +
      "Each is placed at the size and with the alt text listed; set `widthPx` or `heightPx` on the block to resize it " +
      "(the other follows the image's shape), or `alt` to describe it differently. " +
      "Any other `image:` name is refused with `image_unknown`.\n\n" +
      lines
    );
  }

  /** The image `src` description in `Agent.jsonSchema()`. Identical to PHP's. */
  srcDescription(): string {
    return (
      "PNG or JPEG data URL, or a registered image by name: " +
      this.names()
        .map((n) => IMAGE_PREFIX + n)
        .join(", ") +
      ". " +
      "Use a registered name rather than inventing an image; any other image: name is refused (image_unknown)."
    );
  }

  /** The image `src` pattern: a data URL, or exactly one of the registered names. */
  srcPattern(): string {
    return "^(data:image/(png|jpe?g);base64,|" + IMAGE_PREFIX + "(" + this.names().join("|") + ")$)";
  }

  /** The validation error for `image:<name>` naming nothing registered. Mirrors PHP. */
  static unknownError(path: string, name: string, names: string[]): ValidationError {
    return {
      path,
      message:
        `no image named "${name}" is registered` +
        (names.length === 0
          ? "; none are, so use a PNG or JPEG data URL"
          : "; the registered ones are " + names.map((n) => IMAGE_PREFIX + n).join(", ")),
      code: "image_unknown",
    };
  }

  private walk(node: unknown, path: string): unknown {
    if (Array.isArray(node)) return node.map((v, i) => this.walk(v, `${path}/${i}`));
    if (!isPlainObject(node) || isBytes(node)) return node;
    const o = node as Any;
    if (o.type === "image" && typeof o.src === "string" && o.src.startsWith(IMAGE_PREFIX)) {
      const name = o.src.slice(IMAGE_PREFIX.length);
      const image = this.images.get(name);
      if (image === undefined) {
        throw new SchemaException("Doc names an image that is not registered.", [
          NamedImages.unknownError(`${path}/src`, name, this.names()),
        ]);
      }
      const out: Any = { ...o, src: `data:image/${image.type};base64,${base64Encode(image.bytes)}` };
      if (!has(o, "widthPx") && !has(o, "heightPx")) {
        if (image.widthPx !== undefined) out.widthPx = image.widthPx;
        if (image.heightPx !== undefined) out.heightPx = image.heightPx;
      }
      if (!has(o, "alt") && image.alt !== undefined) out.alt = image.alt;
      return out;
    }
    const out: Any = {};
    for (const [key, value] of Object.entries(o)) {
      // A chart's ECharts option is not document content.
      out[key] = key === "option" || value === null || typeof value !== "object" ? value : this.walk(value, `${path}/${key}`);
    }
    return out;
  }

  private static limits(limits: unknown): Limits {
    if (limits === undefined || limits === null) return { ...DEFAULT_IMAGE_LIMITS };
    const emptyArray = Array.isArray(limits) && limits.length === 0;
    if (!emptyArray && !isPlainObject(limits)) {
      throw new ImageRefused("image_limits_invalid", "imageLimits must be an object of {maxBytes?, maxWidthPx?, maxHeightPx?, maxImages?}.");
    }
    const out: Limits = { ...DEFAULT_IMAGE_LIMITS };
    for (const [key, value] of Object.entries(emptyArray ? {} : (limits as Record<string, unknown>))) {
      if (!has(DEFAULT_IMAGE_LIMITS, key)) {
        throw new ImageRefused("image_limits_invalid", "imageLimits takes maxBytes, maxWidthPx, maxHeightPx and maxImages only.");
      }
      if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
        throw new ImageRefused("image_limits_invalid", `imageLimits.${key} must be a positive integer.`);
      }
      out[key as keyof Limits] = value;
    }
    return out;
  }

  /** One image, checked: its options, its bytes, its type, its pixel size. */
  private static register(name: string, input: unknown, limits: Limits): Registered {
    const options: { widthPx?: number; heightPx?: number; alt?: string } = {};
    let image: unknown = input;
    if (isPlainObject(image) && !isBytes(image)) {
      const o = image as Any;
      for (const key of Object.keys(o)) {
        if (!IMAGE_KEYS.includes(key)) {
          throw new ImageRefused("image_option_invalid", `image ${name} takes data, widthPx, heightPx and alt only.`, name);
        }
      }
      if (typeof o.data !== "string" && !isBytes(o.data)) {
        throw new ImageRefused("image_data_invalid", `image ${name} needs its data: PNG or JPEG bytes, or a data URL.`, name);
      }
      for (const dim of ["widthPx", "heightPx"] as const) {
        if (has(o, dim)) {
          if (typeof o[dim] !== "number" || !Number.isInteger(o[dim]) || o[dim] < 1) {
            throw new ImageRefused("image_option_invalid", `image ${name} ${dim} must be a positive integer.`, name);
          }
          options[dim] = o[dim];
        }
      }
      if (has(o, "alt")) {
        if (typeof o.alt !== "string") {
          throw new ImageRefused("image_option_invalid", `image ${name} alt must be text.`, name);
        }
        options.alt = o.alt;
      }
      image = o.data;
    }
    if (typeof image !== "string" && !isBytes(image)) {
      throw new ImageRefused(
        "image_data_invalid",
        `image ${name} must be PNG or JPEG bytes, a data URL, or {data, widthPx?, heightPx?, alt?}.`,
        name,
      );
    }

    const [declared, bytes] = NamedImages.bytes(name, image, limits.maxBytes);
    const type = startsWith(bytes, PNG_SIGNATURE) ? "png" : startsWith(bytes, JPEG_SIGNATURE) ? "jpeg" : null;
    if (type === null || (declared !== null && declared !== type)) {
      throw new ImageRefused(
        "image_type_unsupported",
        `image ${name} must be a PNG or JPEG` + (type !== null ? " whose bytes match its data URL type." : "."),
        name,
      );
    }
    const size = type === "png" ? pngDimensions(bytes) : jpegSize(bytes);
    if (size === null) {
      throw new ImageRefused("image_dimensions_unreadable", `image ${name}: its pixel size cannot be read.`, name);
    }
    if (size.width > limits.maxWidthPx || size.height > limits.maxHeightPx) {
      throw new ImageRefused(
        "image_dimensions_too_large",
        `image ${name} is ${size.width} x ${size.height} px; at most ${limits.maxWidthPx} x ${limits.maxHeightPx} (imageLimits.maxWidthPx / maxHeightPx).`,
        name,
      );
    }
    return { type, bytes, width: size.width, height: size.height, ...options };
  }

  /**
   * The bytes, and the type a data URL declares (null for raw bytes). A data URL is
   * measured before it is decoded, so an oversized one costs nothing.
   */
  private static bytes(name: string, image: string | Uint8Array | ArrayBuffer, maxBytes: number): ["png" | "jpeg" | null, Uint8Array] {
    const tooLarge = (): ImageRefused =>
      new ImageRefused("image_too_large", `image ${name} is over ${maxBytes} bytes (imageLimits.maxBytes).`, name);

    if (typeof image !== "string") {
      const bytes = toU8(image);
      if (bytes.length === 0) throw new ImageRefused("image_data_invalid", `image ${name} is empty.`, name);
      if (bytes.length > maxBytes) throw tooLarge();
      return [null, bytes];
    }
    if (!image.startsWith("data:")) {
      // PHP takes any other string as raw bytes, and those are refused below as
      // not PNG or JPEG; a JavaScript string never holds raw bytes, so the same
      // codes come straight from here.
      if (image === "") throw new ImageRefused("image_data_invalid", `image ${name} is empty.`, name);
      if (image.length > maxBytes) throw tooLarge();
      throw new ImageRefused("image_type_unsupported", `image ${name} must be a PNG or JPEG: bytes, or a data URL.`, name);
    }
    const m = /^data:image\/(png|jpe?g);base64,([\s\S]*)$/.exec(image);
    if (m === null) {
      throw new ImageRefused("image_type_unsupported", `image ${name} must be a base64 PNG or JPEG data URL.`, name);
    }
    const payload = m[2]!;
    if (payload === "" || payload.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(payload)) {
      throw new ImageRefused("image_data_invalid", `image ${name} is not valid base64.`, name);
    }
    const padding = payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0;
    if ((payload.length / 4) * 3 - padding > maxBytes) throw tooLarge();
    const bytes = decodeBase64(payload);
    if (bytes.length === 0) throw new ImageRefused("image_data_invalid", `image ${name} is not valid base64.`, name);
    return [m[1] === "png" ? "png" : "jpeg", bytes];
  }
}
