// Run: 編集済みの VBA マクロ(Sub/Function)を Office 上で実行する。
// 既定では実行前に対象モジュール（またはプロジェクト）を Push し、保存済みコードを実行する。
import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs/promises";
import { getConfig, bridgeOptions } from "../config";
import { invoke } from "../bridge/powershell";
import { scriptFor } from "../bridge/office";
import { RunPayload, RunResult } from "../types";
import { ComponentRecord, readManifest } from "../model/manifest";
import { parseProcedures, runnableMacros } from "../model/procedures";
import { newRecordForFile, runPush } from "./push";
import {
  ProjectEntry,
  findProjectContainingPath,
  findProjectForFile,
  handleBridgeError,
  pickProject,
} from "./shared";

/** コマンドに渡りうる引数（CodeLens / エディタタイトル / ツリー / パレット）。 */
interface RunArg {
  file?: string;
  macro?: string;
  module?: string;
  projectDir?: string;
}

interface MacroCandidate {
  name: string;
  module: string;
  kind: "Sub" | "Function";
}

function normalizeArg(arg: unknown): RunArg {
  if (!arg) return {};
  if (arg instanceof vscode.Uri) return { file: arg.fsPath };
  if (typeof arg === "string") return { file: arg };
  if (typeof arg === "object") {
    const a = arg as Record<string, unknown>;
    const file =
      typeof a.file === "string"
        ? a.file
        : a.resourceUri instanceof vscode.Uri
        ? a.resourceUri.fsPath
        : undefined;
    return {
      file,
      macro: typeof a.macro === "string" ? a.macro : undefined,
      module: typeof a.module === "string" ? a.module : undefined,
      projectDir: typeof a.projectDir === "string" ? a.projectDir : undefined,
    };
  }
  return {};
}

export async function runCommand(arg?: unknown): Promise<void> {
  const a = normalizeArg(arg);

  // 対象プロジェクトと Push スコープを決める（呼び出し元で挙動を変える）。
  // - ツリーのプロジェクト項目                      → プロジェクト全体
  // - 明示的なファイル(CodeLens / エディタ右上 / エクスプローラ) → そのファイル単体
  // - コマンドパレット                              → アクティブファイルからプロジェクトを推定し、
  //   無ければ選択。VBA 以外のファイルを開いていても実行できる（プロジェクト全体スコープ）。
  let entry: ProjectEntry | undefined;
  let fileRec: ComponentRecord | undefined;
  let moduleName = a.module;
  let projectScope = false;

  if (a.projectDir) {
    const manifest = await readManifest(a.projectDir);
    if (manifest) entry = { dir: a.projectDir, manifest };
    projectScope = true;
  } else if (a.file) {
    const found = await findProjectForFile(a.file);
    if (found) {
      entry = found.entry;
      fileRec = found.entry.manifest.components.find(
        (c) => c.name === found.componentName
      );
      moduleName = moduleName ?? found.componentName;
    } else {
      const proj = await findProjectContainingPath(a.file);
      if (!proj) {
        vscode.window.showWarningMessage(
          "Perfect VBA: このファイルは取り込み済みプロジェクトに含まれていません。"
        );
        return;
      }
      entry = proj;
      fileRec = newRecordForFile(proj.manifest, path.basename(a.file));
      moduleName = moduleName ?? fileRec?.name;
    }
  } else {
    const ed = vscode.window.activeTextEditor;
    const hint =
      ed && ed.document.uri.scheme === "file"
        ? await findProjectForFile(ed.document.uri.fsPath)
        : undefined;
    if (hint) {
      entry = hint.entry;
    } else {
      const picked = await pickProject();
      if (!picked) return;
      entry = picked;
    }
    projectScope = true;
  }

  if (!entry) {
    vscode.window.showWarningMessage(
      "Perfect VBA: 実行対象のプロジェクトが見つかりません。先に Pull してください。"
    );
    return;
  }

  // 実行するマクロを確定（引数で指定が無ければ候補から選ばせる）。
  let macro = a.macro;
  let runModule = moduleName;

  if (!macro) {
    const candidates = await discoverMacros(
      entry,
      projectScope ? undefined : fileRec
    );
    if (candidates.length === 0) {
      vscode.window.showWarningMessage(
        "Perfect VBA: 実行できるマクロ（標準モジュール内の引数なし Public Sub/Function）が見つかりませんでした。"
      );
      return;
    }
    const chosen =
      candidates.length === 1 && !projectScope
        ? candidates[0]
        : await pickMacro(candidates);
    if (!chosen) return;
    macro = chosen.name;
    runModule = chosen.module;
    if (!fileRec) {
      const owner = chosen.module;
      fileRec = entry.manifest.components.find((c) => c.name === owner);
    }
  }
  if (!macro) return;

  const cfg = getConfig();

  // 実行前 Push（保存済みコードを実行するため）。失敗したら中断。
  if (cfg.pushBeforeRun) {
    await saveDirtyDocs(entry.dir, projectScope ? undefined : fileRec?.file);
    const records = projectScope
      ? entry.manifest.components
      : fileRec
      ? [fileRec]
      : [];
    if (records.length > 0) {
      const ok = await runPush(entry.dir, entry.manifest, records, {
        silent: true,
        save: true,
      });
      if (!ok) return;
    }
  }

  // 実行
  const app = entry.manifest.app;
  const source = entry.manifest.source;
  try {
    const { data } = await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: `Perfect VBA: ${macro} を実行中...`,
        cancellable: true,
      },
      (_progress, token) =>
        invoke<RunResult>(
          scriptFor(app, "run"),
          {
            path: source,
            headless: cfg.headless,
            macro,
            module: runModule,
          } as RunPayload,
          bridgeOptions(cfg),
          token
        )
    );
    const rv = data.returnValue;
    if (rv !== null && rv !== undefined && `${rv}`.length > 0) {
      vscode.window.showInformationMessage(
        `Perfect VBA: ${macro} を実行しました（戻り値: ${rv}）`
      );
    } else {
      vscode.window.showInformationMessage(
        `Perfect VBA: ${macro} を実行しました。`
      );
    }
  } catch (err) {
    await handleBridgeError(err, app);
  }
}

/** プロジェクト（または単一モジュール）の標準モジュールから実行可能マクロを集める。 */
async function discoverMacros(
  entry: ProjectEntry,
  only?: ComponentRecord
): Promise<MacroCandidate[]> {
  const recs = only ? [only] : entry.manifest.components;
  const out: MacroCandidate[] = [];
  for (const rec of recs) {
    if (rec.kind !== "std") continue; // 名前実行できるのは標準モジュールのみ
    let text: string;
    try {
      text = await fs.readFile(path.join(entry.dir, rec.file), "utf8");
    } catch {
      continue;
    }
    for (const proc of runnableMacros(parseProcedures(text))) {
      out.push({ name: proc.name, module: rec.name, kind: proc.kind });
    }
  }
  return out;
}

async function pickMacro(
  candidates: MacroCandidate[]
): Promise<MacroCandidate | undefined> {
  const pick = await vscode.window.showQuickPick(
    candidates
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((c) => ({
        label: c.name,
        description: c.module,
        detail: c.kind,
        candidate: c,
      })),
    { placeHolder: "実行するマクロを選択" }
  );
  return pick?.candidate;
}

/** 実行前に、対象（file 指定時はそのファイル、未指定時は dir 配下全部）の未保存を保存する。 */
async function saveDirtyDocs(dir: string, file?: string): Promise<void> {
  const root = path.resolve(dir);
  const target = file ? path.resolve(path.join(dir, file)) : undefined;
  for (const doc of vscode.workspace.textDocuments) {
    if (!doc.isDirty || doc.uri.scheme !== "file") continue;
    const p = path.resolve(doc.uri.fsPath);
    if (target) {
      if (p === target) await doc.save();
    } else if (p.startsWith(root)) {
      await doc.save();
    }
  }
}
