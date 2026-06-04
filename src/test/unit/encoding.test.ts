import * as assert from "node:assert/strict";
import {
  resolveCodepage,
  decodeBytes,
  encodeText,
  encodeForWrite,
  detectBom,
  sniffEncoding,
} from "../../encoding";

// 日本語コメントとエクスポートヘッダを含む、典型的な .bas の中身
const SAMPLE_BAS = [
  'Attribute VB_Name = "Module1"',
  "Option Explicit",
  "",
  "Sub こんにちは()",
  "    ' 日本語のコメント：全角文字を含む",
  '    MsgBox "テスト①②③"',
  "End Sub",
].join("\r\n");

describe("resolveCodepage", () => {
  it("shift_jis -> cp932", () => {
    assert.equal(resolveCodepage("shift_jis"), "cp932");
  });
  it("utf-8 -> utf8", () => {
    assert.equal(resolveCodepage("utf-8"), "utf8");
  });
  it("auto + acp 932 -> cp932", () => {
    assert.equal(resolveCodepage("auto", 932), "cp932");
  });
  it("auto + acp 1252 -> cp1252", () => {
    assert.equal(resolveCodepage("auto", 1252), "cp1252");
  });
  it("auto + unknown acp -> フォールバック cp932", () => {
    assert.equal(resolveCodepage("auto", undefined), "cp932");
  });
});

describe("cp932 ラウンドトリップ", () => {
  it("日本語コメント＋ヘッダが往復で一致する", () => {
    const bytes = encodeText(SAMPLE_BAS, "cp932");
    const back = decodeBytes(bytes, "cp932");
    assert.equal(back, SAMPLE_BAS);
  });

  it("Attribute VB_Name ヘッダ行が保持される", () => {
    const bytes = encodeText(SAMPLE_BAS, "cp932");
    const back = decodeBytes(bytes, "cp932");
    assert.ok(back.startsWith('Attribute VB_Name = "Module1"'));
  });

  it("cp932 でエンコードしたバイトは UTF-8 とは異なる（実際に変換している）", () => {
    const sjis = encodeText(SAMPLE_BAS, "cp932");
    const utf8 = Buffer.from(SAMPLE_BAS, "utf8");
    assert.notEqual(sjis.length, utf8.length);
  });
});

describe("BOM 検出とデコード", () => {
  it("UTF-8 BOM を検出して剥がす", () => {
    const buf = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from("あ", "utf8")]);
    assert.deepEqual(detectBom(buf), { enc: "utf8", length: 3 });
    assert.equal(decodeBytes(buf, "cp932"), "あ");
  });

  it("UTF-16LE BOM を検出してデコードする", () => {
    const body = Buffer.from("テスト", "utf16le");
    const buf = Buffer.concat([Buffer.from([0xff, 0xfe]), body]);
    assert.deepEqual(detectBom(buf), { enc: "utf16le", length: 2 });
    assert.equal(decodeBytes(buf, "cp932"), "テスト");
  });

  it("BOM 無しは fallback を使う", () => {
    const buf = encodeText("ABC", "cp932");
    assert.equal(sniffEncoding(buf, "cp932"), "cp932");
  });
});

describe("encodeForWrite", () => {
  it("utf16le では BOM を付与する", () => {
    const buf = encodeForWrite("X", "utf16le");
    assert.equal(buf[0], 0xff);
    assert.equal(buf[1], 0xfe);
  });
  it("cp932 では BOM を付与しない", () => {
    const buf = encodeForWrite("X", "cp932");
    assert.equal(buf[0], "X".charCodeAt(0));
  });
  it("utf16le は encodeForWrite -> decodeBytes で往復一致する", () => {
    const enc = encodeForWrite(SAMPLE_BAS, "utf16le");
    assert.equal(decodeBytes(enc, "cp932"), SAMPLE_BAS);
  });
});
