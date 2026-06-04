import * as assert from "node:assert/strict";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { randomUUID } from "node:crypto";
import {
  Manifest,
  MANIFEST_VERSION,
  projectFolderName,
  readManifest,
  sha256,
  writeManifest,
} from "../../model/manifest";
import { componentFileName, importExt, sanitizeName } from "../../model/layout";

describe("sha256", () => {
  it("同じ入力は同じハッシュ", () => {
    assert.equal(sha256("abc"), sha256("abc"));
  });
  it("異なる入力は異なるハッシュ", () => {
    assert.notEqual(sha256("abc"), sha256("abd"));
  });
});

describe("projectFolderName", () => {
  it("ベース名を返す", () => {
    assert.equal(projectFolderName("C:/work/Book1.xlsm"), "Book1.xlsm");
  });
});

describe("componentFileName / importExt", () => {
  it("std -> .bas", () => {
    assert.deepEqual(componentFileName("Module1", "std"), { file: "Module1.bas" });
    assert.equal(importExt("std"), ".bas");
  });
  it("class -> .cls", () => {
    assert.deepEqual(componentFileName("clsThing", "class"), { file: "clsThing.cls" });
  });
  it("form -> .frm + .frx", () => {
    assert.deepEqual(componentFileName("UserForm1", "form"), {
      file: "UserForm1.frm",
      frxFile: "UserForm1.frx",
    });
  });
  it("document -> .cls", () => {
    assert.deepEqual(componentFileName("ThisWorkbook", "document"), {
      file: "ThisWorkbook.cls",
    });
  });
  it("accForm -> Form_<name>.txt", () => {
    assert.deepEqual(componentFileName("frmMain", "accForm"), { file: "Form_frmMain.txt" });
    assert.equal(importExt("accForm"), ".txt");
  });
  it("不正文字をサニタイズする", () => {
    assert.equal(sanitizeName('a/b:c*'), "a_b_c_");
  });
});

describe("manifest 読み書き", () => {
  it("書き込んだ内容を読み戻せる", async () => {
    const dir = path.join(os.tmpdir(), "perfectvba-test-" + randomUUID());
    try {
      const manifest: Manifest = {
        version: MANIFEST_VERSION,
        app: "excel",
        source: "C:/work/Book1.xlsm",
        projectName: "VBAProject",
        encodingSetting: "auto",
        pulledAt: "2026-06-02T00:00:00.000Z",
        components: [
          { name: "Module1", kind: "std", file: "Module1.bas", enc: "cp932", hash: sha256("x") },
        ],
      };
      await writeManifest(dir, manifest);
      const back = await readManifest(dir);
      assert.deepEqual(back, manifest);
    } finally {
      await fs.rm(dir, { recursive: true, force: true });
    }
  });

  it("存在しないディレクトリは undefined", async () => {
    const back = await readManifest(path.join(os.tmpdir(), "nope-" + randomUUID()));
    assert.equal(back, undefined);
  });
});
