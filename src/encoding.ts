// エンコーディング変換。VBE の Export は日本語 Windows では Shift_JIS(cp932)、
// VSCode は UTF-8。ここで cp932 等 ↔ UTF-8 を吸収する。
// Node 標準では cp932 をデコードできないため iconv-lite を使う。
import * as iconv from "iconv-lite";

/** 設定値。 */
export type EncodingSetting = "auto" | "shift_jis" | "utf-8";

/** ファイルに実際に適用するエンコーディングの解決済み名（iconv-lite の名前 or 特殊値）。 */
export type ResolvedEncoding = string; // 例: "cp932", "utf8", "utf16le"

/**
 * 設定とシステム ANSI コードページ番号から、byte 種別に使う iconv 名を解決する。
 * - "shift_jis" -> cp932（MS 拡張を含む上位互換）
 * - "utf-8"     -> utf8
 * - "auto"      -> 'cp' + acp（存在すれば）。日本語環境では acp=932 -> cp932
 */
export function resolveCodepage(
  setting: EncodingSetting,
  acp?: number
): ResolvedEncoding {
  if (setting === "shift_jis") return "cp932";
  if (setting === "utf-8") return "utf8";
  // auto
  if (acp && acp > 0) {
    const candidate = `cp${acp}`;
    if (iconv.encodingExists(candidate)) return candidate;
  }
  // ACP 不明時のフォールバック（日本語環境想定）
  return "cp932";
}

/** バイト列を指定エンコーディングで UTF-8 文字列にデコードする（BOM を考慮）。 */
export function decodeBytes(buf: Buffer, enc: ResolvedEncoding): string {
  const bom = detectBom(buf);
  if (bom) {
    return iconv.decode(buf.subarray(bom.length), bom.enc);
  }
  return iconv.decode(buf, enc);
}

/** UTF-8 文字列を指定エンコーディングのバイト列にエンコードする。 */
export function encodeText(text: string, enc: ResolvedEncoding): Buffer {
  return iconv.encode(text, enc);
}

/**
 * ファイル書き込み用にエンコードする。UTF-16 系は BOM を付与する
 * （Access SaveAsText/LoadFromText の往復で BOM を保つため）。
 */
export function encodeForWrite(text: string, enc: ResolvedEncoding): Buffer {
  if (enc === "utf16le") {
    return Buffer.concat([Buffer.from([0xff, 0xfe]), iconv.encode(text, "utf16le")]);
  }
  if (enc === "utf16be") {
    return Buffer.concat([Buffer.from([0xfe, 0xff]), iconv.encode(text, "utf16be")]);
  }
  return iconv.encode(text, enc);
}

/** BOM を検出し、対応する iconv 名と BOM バイト長を返す。無ければ undefined。 */
export function detectBom(
  buf: Buffer
): { enc: ResolvedEncoding; length: number } | undefined {
  if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) {
    return { enc: "utf8", length: 3 };
  }
  if (buf.length >= 2 && buf[0] === 0xff && buf[1] === 0xfe) {
    return { enc: "utf16le", length: 2 };
  }
  if (buf.length >= 2 && buf[0] === 0xfe && buf[1] === 0xff) {
    return { enc: "utf16be", length: 2 };
  }
  return undefined;
}

/**
 * バイト列から最終的に保存に使うエンコーディングを推定する。
 * BOM があればそれを優先、無ければ渡された既定(codepage)を使う。
 * Access の SaveAsText 出力など、エンコーディングが版によって異なるケースに対応。
 */
export function sniffEncoding(buf: Buffer, fallback: ResolvedEncoding): ResolvedEncoding {
  return detectBom(buf)?.enc ?? fallback;
}
