// コマンド間で共有するヘルパ（ワークスペース解決・ファイル選択・プロジェクト列挙・エラー処理）。
import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs/promises";
import { getConfig } from "../config";
import { OFFICE_FILTERS } from "../bridge/office";
import { BridgeError } from "../bridge/powershell";
import { AppKind } from "../types";
import { Manifest, MANIFEST_NAME, readManifest } from "../model/manifest";

export function firstWorkspaceFolder(): vscode.WorkspaceFolder | undefined {
  return vscode.workspace.workspaceFolders?.[0];
}

export function projectsRootDir(): string | undefined {
  const ws = firstWorkspaceFolder();
  if (!ws) return undefined;
  return path.join(ws.uri.fsPath, getConfig().workspaceRoot);
}

export async function pickOfficeFile(): Promise<string | undefined> {
  const picked = await vscode.window.showOpenDialog({
    canSelectMany: false,
    openLabel: "取り込む",
    filters: OFFICE_FILTERS,
  });
  return picked?.[0]?.fsPath;
}

export interface ProjectEntry {
  dir: string;
  manifest: Manifest;
}

/** ワークスペースルート配下の取り込み済みプロジェクトを列挙する。 */
export async function listProjects(): Promise<ProjectEntry[]> {
  const root = projectsRootDir();
  if (!root) return [];
  let entries: import("fs").Dirent[];
  try {
    entries = await fs.readdir(root, { withFileTypes: true });
  } catch {
    return [];
  }
  const result: ProjectEntry[] = [];
  for (const e of entries) {
    if (!e.isDirectory()) continue;
    const dir = path.join(root, e.name);
    const manifest = await readManifest(dir);
    if (manifest) result.push({ dir, manifest });
  }
  return result;
}

/** 任意のファイルパスを含むプロジェクト（manifest にそのファイルが載っているもの）を探す。 */
export async function findProjectForFile(
  filePath: string
): Promise<{ entry: ProjectEntry; componentName: string } | undefined> {
  const projects = await listProjects();
  const target = path.resolve(filePath);
  for (const entry of projects) {
    for (const comp of entry.manifest.components) {
      if (path.resolve(path.join(entry.dir, comp.file)) === target) {
        return { entry, componentName: comp.name };
      }
    }
  }
  return undefined;
}

/** ファイルパスを配下に含むプロジェクト（manifest 未登録でも可）を探す。 */
export async function findProjectContainingPath(
  filePath: string
): Promise<ProjectEntry | undefined> {
  const projects = await listProjects();
  const target = path.resolve(filePath);
  for (const entry of projects) {
    const root = path.resolve(entry.dir);
    const rel = path.relative(root, target);
    if (rel && !rel.startsWith("..") && !path.isAbsolute(rel)) return entry;
  }
  return undefined;
}

/** クイックピックでプロジェクトを 1 つ選ばせる。 */
export async function pickProject(): Promise<ProjectEntry | undefined> {
  const projects = await listProjects();
  if (projects.length === 0) {
    vscode.window.showWarningMessage(
      "取り込み済みの VBA プロジェクトがありません。先に Pull してください。"
    );
    return undefined;
  }
  if (projects.length === 1) return projects[0];
  const pick = await vscode.window.showQuickPick(
    projects.map((p) => ({
      label: p.manifest.projectName,
      description: path.basename(p.manifest.source),
      detail: p.manifest.source,
      entry: p,
    })),
    { placeHolder: "Push するプロジェクトを選択" }
  );
  return pick?.entry;
}

export async function handleBridgeError(err: unknown, app: AppKind): Promise<void> {
  if (err instanceof BridgeError && err.code === "TRUST_DISABLED") {
    const enable = "信頼設定を有効化";
    const choice = await vscode.window.showErrorMessage(
      "VBA プロジェクトへのプログラムによるアクセスが信頼されていません。" +
        "Excel/Access の [トラスト センター] → [マクロの設定] → " +
        "[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] を有効にしてください。",
      enable
    );
    if (choice === enable) {
      await vscode.commands.executeCommand("perfectVba.enableTrust", app);
    }
    return;
  }
  if (err instanceof BridgeError && err.code === "EXCLUSIVE_LOCKED") {
    vscode.window.showErrorMessage(
      "Perfect VBA: データベースを排他オープンできませんでした。対象を開いている他のプロセスを閉じてから再試行してください。"
    );
    return;
  }
  const msg = err instanceof Error ? err.message : String(err);
  vscode.window.showErrorMessage(`Perfect VBA: ${msg}`);
}

export async function safeUnlink(p: string): Promise<void> {
  try {
    await fs.unlink(p);
  } catch {
    /* ignore */
  }
}

export { MANIFEST_NAME };
