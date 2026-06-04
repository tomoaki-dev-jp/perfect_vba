import * as assert from "node:assert/strict";
import {
  parseProcedures,
  runnableMacros,
} from "../../model/procedures";

describe("parseProcedures", () => {
  it("基本的な Sub / Function を抽出する", () => {
    const src = [
      "Sub Foo()",
      "End Sub",
      "Public Function Bar() As Long",
      "End Function",
    ].join("\r\n");
    const procs = parseProcedures(src);
    assert.equal(procs.length, 2);
    assert.deepEqual(
      procs.map((p) => [p.name, p.kind, p.visibility, p.hasRequiredParams]),
      [
        ["Foo", "Sub", "Default", false],
        ["Bar", "Function", "Public", false],
      ]
    );
    assert.equal(procs[0].line, 0);
    assert.equal(procs[1].line, 2);
  });

  it("可視性(Public/Private/Friend)と Static を判定する", () => {
    const src = [
      "Private Sub Secret()",
      "Friend Sub Buddy()",
      "Public Static Sub Keeper()",
    ].join("\n");
    const procs = parseProcedures(src);
    assert.deepEqual(
      procs.map((p) => [p.name, p.visibility]),
      [
        ["Secret", "Private"],
        ["Buddy", "Friend"],
        ["Keeper", "Public"],
      ]
    );
  });

  it("必須引数の有無を判定する（Optional / ParamArray は必須でない）", () => {
    const src = [
      "Sub NeedsArg(ByVal x As Long)",
      "Sub OptOnly(Optional ByVal y As Long = 1)",
      "Sub Variadic(ParamArray args() As Variant)",
      "Sub NoArg()",
    ].join("\n");
    const map = Object.fromEntries(
      parseProcedures(src).map((p) => [p.name, p.hasRequiredParams])
    );
    assert.deepEqual(map, {
      NeedsArg: true,
      OptOnly: false,
      Variadic: false,
      NoArg: false,
    });
  });

  it("行継続( _ )をまたぐ宣言を 1 つの手続きとして扱う", () => {
    const src = [
      "Public Sub Wide(ByVal a As Long, _",
      "                ByVal b As Long)",
      "End Sub",
    ].join("\r\n");
    const procs = parseProcedures(src);
    assert.equal(procs.length, 1);
    assert.equal(procs[0].name, "Wide");
    assert.equal(procs[0].hasRequiredParams, true);
    assert.equal(procs[0].line, 0);
  });

  it("Declare / Property / コメント / End Sub は手続きとして拾わない", () => {
    const src = [
      "' Sub Commented()",
      "Rem Sub AlsoCommented()",
      'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)',
      "Property Get Value() As Long",
      "End Property",
      "Sub Real()",
      "    Exit Sub",
      "End Sub",
    ].join("\n");
    const procs = parseProcedures(src);
    assert.deepEqual(
      procs.map((p) => p.name),
      ["Real"]
    );
  });
});

describe("runnableMacros", () => {
  it("Private と必須引数ありを除外する", () => {
    const src = [
      "Sub RunMe()",
      "End Sub",
      "Public Function Calc() As Long",
      "End Function",
      "Private Sub Hidden()",
      "End Sub",
      "Sub WithArg(ByVal n As Long)",
      "End Sub",
    ].join("\n");
    const runnable = runnableMacros(parseProcedures(src)).map((p) => p.name);
    assert.deepEqual(runnable, ["RunMe", "Calc"]);
  });
});
