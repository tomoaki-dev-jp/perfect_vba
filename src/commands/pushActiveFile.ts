// 現在アクティブな .bas/.cls/.frm/.txt を、その所属プロジェクトへ単体 Push する。
import * as vscode from "vscode";
import * as path from "path";
import { findProjectContainingPath, findProjectForFile } from "./shared";
import { newRecordForFile, runPush } from "./push";

export async function pushActiveFileCommand(): Promise<void> {
  const editor = vscode.window.activeTextEditor;
  if (!editor) {
    vscode.window.showWarningMessage("Perfect VBA: アクティブなエディタがありません。");
    return;
  }
  if (editor.document.isDirty) {
    await editor.document.save();
  }
  const file = editor.document.uri.fsPath;

  // 既に取り込み済みのファイル。
  const found = await findProjectForFile(file);
  if (found) {
    const rec = found.entry.manifest.components.find(
      (c) => c.name === found.componentName
    );
    if (!rec) return;
    await runPush(found.entry.dir, found.entry.manifest, [rec]);
    return;
  }

  // manifest 未登録でも、プロジェクトフォルダ配下の .bas/.cls/.frm なら新規として Push。
  const proj = await findProjectContainingPath(file);
  if (proj) {
    const rec = newRecordForFile(proj.manifest, path.basename(file));
    if (rec) {
      await runPush(proj.dir, proj.manifest, [rec]);
      return;
    }
    vscode.window.showWarningMessage(
      "Perfect VBA: このファイル形式は Push 対象外です（.bas / .cls / .frm のみ）。"
    );
    return;
  }

  vscode.window.showWarningMessage(
    "Perfect VBA: このファイルは取り込み済みプロジェクトに含まれていません。"
  );
}
