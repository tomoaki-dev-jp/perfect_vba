// 紐づく Office ファイルを既定アプリで開く。
import * as vscode from "vscode";
import { pickProject } from "./shared";

interface SourceArg {
  manifestSource?: string;
}

export async function openSourceCommand(arg?: unknown): Promise<void> {
  let source: string | undefined;
  if (arg && typeof arg === "object" && "manifestSource" in arg) {
    source = (arg as SourceArg).manifestSource;
  }
  if (!source) {
    const entry = await pickProject();
    source = entry?.manifest.source;
  }
  if (!source) return;

  const ok = await vscode.env.openExternal(vscode.Uri.file(source));
  if (!ok) {
    vscode.window.showErrorMessage(`Perfect VBA: ファイルを開けませんでした: ${source}`);
  }
}
