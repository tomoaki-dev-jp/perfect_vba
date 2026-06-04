// 標準モジュール(.bas)の各「実行可能マクロ」の上に ▶ 実行 ボタン(CodeLens)を出す。
import * as vscode from "vscode";
import * as path from "path";
import { parseProcedures, runnableMacros } from "../model/procedures";
import { projectsRootDir } from "../commands/shared";

export class RunCodeLensProvider implements vscode.CodeLensProvider {
  private readonly _onDidChange = new vscode.EventEmitter<void>();
  readonly onDidChangeCodeLenses = this._onDidChange.event;

  refresh(): void {
    this._onDidChange.fire();
  }

  provideCodeLenses(document: vscode.TextDocument): vscode.CodeLens[] {
    if (document.languageId !== "vba") return [];
    if (path.extname(document.uri.fsPath).toLowerCase() !== ".bas") return [];

    // 取り込み済みプロジェクト(ワークスペースルート配下)のファイルにのみ出す。
    const root = projectsRootDir();
    if (!root) return [];
    const filePath = path.resolve(document.uri.fsPath);
    if (!filePath.startsWith(path.resolve(root))) return [];

    const moduleName = path.basename(
      document.uri.fsPath,
      path.extname(document.uri.fsPath)
    );

    const lenses: vscode.CodeLens[] = [];
    for (const proc of runnableMacros(parseProcedures(document.getText()))) {
      const range = new vscode.Range(proc.line, 0, proc.line, 0);
      lenses.push(
        new vscode.CodeLens(range, {
          title: "$(play) 実行",
          tooltip: `${moduleName}.${proc.name} を Push して実行`,
          command: "perfectVba.run",
          arguments: [
            { file: document.uri.fsPath, macro: proc.name, module: moduleName },
          ],
        })
      );
    }
    return lenses;
  }
}
