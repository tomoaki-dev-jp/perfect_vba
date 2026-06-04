// ステータスバー: アクティブファイルが取り込み済みプロジェクト内なら Push ボタンを表示。
import * as vscode from "vscode";
import { findProjectForFile } from "../commands/shared";

export function createStatusBar(context: vscode.ExtensionContext): vscode.StatusBarItem {
  const item = vscode.window.createStatusBarItem(
    vscode.StatusBarAlignment.Right,
    100
  );
  item.command = "perfectVba.pushActiveFile";

  const update = async (): Promise<void> => {
    const editor = vscode.window.activeTextEditor;
    if (!editor || editor.document.uri.scheme !== "file") {
      item.hide();
      return;
    }
    const found = await findProjectForFile(editor.document.uri.fsPath);
    if (found) {
      item.text = "$(cloud-upload) VBA Push";
      item.tooltip = `${found.entry.manifest.projectName} / ${found.componentName} を ${found.entry.manifest.app === "access" ? "Access" : "Excel"} へ Push`;
      item.show();
    } else {
      item.hide();
    }
  };

  context.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor(() => void update())
  );
  void update();
  return item;
}
