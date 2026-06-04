// 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定を有効化する。
import * as vscode from "vscode";
import { getConfig, bridgeOptions } from "../config";
import { invoke } from "../bridge/powershell";
import { AppKind } from "../types";

export async function enableTrustCommand(app?: AppKind): Promise<void> {
  let target = app;
  if (target !== "excel" && target !== "access") {
    const pick = await vscode.window.showQuickPick(
      [
        { label: "Excel", value: "excel" as AppKind },
        { label: "Access", value: "access" as AppKind },
      ],
      { placeHolder: "信頼設定を有効化するアプリを選択" }
    );
    target = pick?.value;
  }
  if (!target) return;

  const cfg = getConfig();
  try {
    const { data } = await invoke<{ versions: string[]; appName: string }>(
      "enable-trust.ps1",
      { app: target },
      bridgeOptions(cfg)
    );
    const vers = data.versions?.length ? data.versions.join(", ") : "(検出なし)";
    const choice = await vscode.window.showInformationMessage(
      `Perfect VBA: ${data.appName} の信頼設定を有効化しました (バージョン: ${vers})。` +
        `反映には ${data.appName} を一度終了して再起動してください。`,
      "OK"
    );
    void choice;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    vscode.window.showErrorMessage(
      `Perfect VBA: 信頼設定の有効化に失敗しました: ${msg}`
    );
  }
}
