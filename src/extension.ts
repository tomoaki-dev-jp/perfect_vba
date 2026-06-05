import * as vscode from "vscode";
import { pullCommand, refreshCommand } from "./commands/pull";
import { pushCommand } from "./commands/push";
import { pushActiveFileCommand } from "./commands/pushActiveFile";
import { runCommand } from "./commands/run";
import { enableTrustCommand } from "./commands/trust";
import { openSourceCommand } from "./commands/openSource";
import { AppKind } from "./types";
import { PerfectVbaTreeProvider } from "./ui/treeView";
import { createStatusBar } from "./ui/statusBar";
import { RunCodeLensProvider } from "./ui/runLens";
import { AutoPushController, toggleAutoPushCommand } from "./autoPush";

export function activate(context: vscode.ExtensionContext): void {
  const tree = new PerfectVbaTreeProvider();
  const runLens = new RunCodeLensProvider();

  context.subscriptions.push(
    vscode.window.registerTreeDataProvider("perfectVbaProjects", tree),
    vscode.languages.registerCodeLensProvider({ language: "vba" }, runLens),

    vscode.commands.registerCommand("perfectVba.pull", (arg?: vscode.Uri) =>
      pullCommand(arg)
    ),
    vscode.commands.registerCommand("perfectVba.push", (arg?: unknown) =>
      pushCommand(arg)
    ),
    vscode.commands.registerCommand("perfectVba.pushActiveFile", () =>
      pushActiveFileCommand()
    ),
    vscode.commands.registerCommand("perfectVba.run", (arg?: unknown) =>
      runCommand(arg)
    ),
    vscode.commands.registerCommand("perfectVba.refresh", (arg?: unknown) =>
      refreshCommand(arg)
    ),
    vscode.commands.registerCommand("perfectVba.refreshTree", () =>
      tree.refresh()
    ),
    vscode.commands.registerCommand("perfectVba.enableTrust", (app?: AppKind) =>
      enableTrustCommand(app)
    ),
    vscode.commands.registerCommand("perfectVba.openSource", (arg?: unknown) =>
      openSourceCommand(arg)
    ),
    vscode.commands.registerCommand("perfectVba.toggleAutoPush", () =>
      toggleAutoPushCommand()
    ),

    // 編集に追従して ▶ 実行 ボタンを出し直す。
    vscode.workspace.onDidSaveTextDocument(() => runLens.refresh()),
    vscode.window.onDidChangeActiveTextEditor(() => runLens.refresh()),

    createStatusBar(context),

    // 保存時の自動 Push（リアルタイム反映）。既定オフ・設定でオン。
    new AutoPushController()
  );
}

export function deactivate(): void {
  /* no-op */
}
