import * as assert from "node:assert/strict";
import { ComponentRecord } from "../../model/manifest";
import {
  duplicateNamesMessage,
  findDuplicateComponentNames,
  normalizeModuleName,
} from "../../model/validate";

function rec(name: string, file: string, kind: ComponentRecord["kind"] = "std"): ComponentRecord {
  return { name, kind, file, enc: "cp932", hash: "" };
}

describe("normalizeModuleName", () => {
  it("大文字小文字を無視する（VBA 識別子は case-insensitive）", () => {
    assert.equal(normalizeModuleName("Foo"), normalizeModuleName("foo"));
    assert.equal(normalizeModuleName("MODULE1"), "module1");
  });

  it("前後の空白を落とす", () => {
    assert.equal(normalizeModuleName("  Foo  "), "foo");
    assert.equal(normalizeModuleName("Foo "), normalizeModuleName("Foo"));
  });

  it("ユニコード合成の揺れ（濁点の合成/分解）を吸収する", () => {
    const composed = "ガ"; // U+30AC
    const decomposed = "ガ".normalize("NFD"); // U+30AB U+3099
    assert.notEqual(composed, decomposed);
    assert.equal(normalizeModuleName(composed), normalizeModuleName(decomposed));
  });

  it("全角/半角は別名として保持する（NFKC はしない）", () => {
    assert.notEqual(normalizeModuleName("ＡＢＣ"), normalizeModuleName("ABC"));
  });
});

describe("findDuplicateComponentNames", () => {
  it("重複が無ければ空", () => {
    const groups = findDuplicateComponentNames([
      rec("Foo", "Foo.bas"),
      rec("Bar", "Bar.bas"),
    ]);
    assert.equal(groups.length, 0);
  });

  it("同名・別ファイル（.bas と .cls）を検出する", () => {
    const groups = findDuplicateComponentNames([
      rec("Foo", "Foo.bas", "std"),
      rec("Foo", "Foo.cls", "class"),
    ]);
    assert.equal(groups.length, 1);
    assert.equal(groups[0].records.length, 2);
    assert.deepEqual(
      groups[0].records.map((r) => r.file),
      ["Foo.bas", "Foo.cls"]
    );
  });

  it("大小・空白だけ違う名前も重複として検出する", () => {
    const groups = findDuplicateComponentNames([
      rec("Module1", "Module1.bas"),
      rec("module1 ", "module1 .bas"),
    ]);
    assert.equal(groups.length, 1);
    assert.equal(groups[0].records.length, 2);
  });

  it("複数の重複グループをまとめて返す", () => {
    const groups = findDuplicateComponentNames([
      rec("Foo", "Foo.bas"),
      rec("Foo", "Foo.cls"),
      rec("Bar", "Bar.bas"),
      rec("bar", "bar.cls"),
      rec("Unique", "Unique.bas"),
    ]);
    assert.equal(groups.length, 2);
  });
});

describe("duplicateNamesMessage", () => {
  it("衝突したファイル名を含む案内文を作る", () => {
    const groups = findDuplicateComponentNames([
      rec("Foo", "Foo.bas"),
      rec("Foo", "Foo.cls"),
    ]);
    const msg = duplicateNamesMessage(groups);
    assert.match(msg, /モジュール名が重複/);
    assert.match(msg, /Foo\.bas/);
    assert.match(msg, /Foo\.cls/);
    assert.match(msg, /プロシージャが見つかりません/);
  });
});
