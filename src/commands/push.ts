// Push: ワークスペースの VBA テキストファイル群 → Office ファイルへ書き戻し。
import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs/promises";
import { getConfig, bridgeOptions } from "../config";
import { invoke } from "../bridge/powershell";
import { scriptFor } from "../bridge/office";
import { ComponentKind, PushComponent, PushPayload } from "../types";
import { encodeForWrite } from "../encoding";
import {
  ComponentRecord,
  Manifest,
  readManifest,
  sha256,
  writeManifest,
} from "../model/manifest";
import { importExt } from "../model/layout";
import {
  duplicateNamesMessage,
  findDuplicateComponentNames,
} from "../model/validate";
import { handleBridgeError, pickProject } from "./shared";

/** ツリー項目などから渡される、プロジェクトフォルダを指す引数。 */
interface ProjectArg {
  projectDir?: string;
}

function resolveDirFromArg(arg: unknown): string | undefined {
  if (typeof arg === "string") return arg;
  if (arg && typeof arg === "object" && "projectDir" in arg) {
    return (arg as ProjectArg).projectDir;
  }
  return undefined;
}

export async function pushCommand(arg?: unknown): Promise<void> {
  let dir = resolveDirFromArg(arg);
  if (!dir) {
    const entry = await pickProject();
    if (!entry) return;
    dir = entry.dir;
  }
  const manifest = await readManifest(dir);
  if (!manifest) {
    vscode.window.showErrorMessage(
      "Perfect VBA: このフォルダに .perfect-vba.json が見つかりません。"
    );
    return;
  }
  // manifest 登録済み + VSCode 側で新規追加された未登録ファイルの両方を Push する。
  const untracked = await discoverUntrackedComponents(dir, manifest);
  await runPush(dir, manifest, [...manifest.components, ...untracked]);
}

/** 拡張子 → コンポーネント種別。新規ファイルの種別推定に使う。 */
const COMPONENT_EXT: { [ext: string]: ComponentKind } = {
  ".bas": "std",
  ".cls": "class",
  ".frm": "form",
};

/** 単一ファイル名から新規コンポーネントのレコードを作る（対象外拡張子なら undefined）。 */
export function newRecordForFile(
  manifest: Manifest,
  fileName: string
): ComponentRecord | undefined {
  const kind = COMPONENT_EXT[path.extname(fileName).toLowerCase()];
  if (!kind) return undefined;
  const name = path.basename(fileName, path.extname(fileName));
  // byte 種別の既定エンコーディングは、プロジェクトの既存レコードの値を踏襲する。
  const enc =
    manifest.components.find((c) => c.enc && c.enc !== "document")?.enc ??
    "cp932";
  const rec: ComponentRecord = { name, kind, file: fileName, enc, hash: "" };
  if (kind === "form") rec.frxFile = `${name}.frx`;
  return rec;
}

/** dir 内の、manifest 未登録の .bas/.cls/.frm を新規コンポーネント候補として返す。 */
async function discoverUntrackedComponents(
  dir: string,
  manifest: Manifest
): Promise<ComponentRecord[]> {
  const tracked = new Set<string>();
  for (const c of manifest.components) {
    tracked.add(c.file.toLowerCase());
    if (c.frxFile) tracked.add(c.frxFile.toLowerCase());
  }
  let entries: import("fs").Dirent[];
  try {
    entries = await fs.readdir(dir, { withFileTypes: true });
  } catch {
    return [];
  }
  const result: ComponentRecord[] = [];
  for (const e of entries) {
    if (!e.isFile() || tracked.has(e.name.toLowerCase())) continue;
    const rec = newRecordForFile(manifest, e.name);
    if (rec) result.push(rec);
  }
  return result;
}

export type PushProgressKind = "notification" | "window" | "none";

export interface RunPushOptions {
  /** 成功トーストを抑制する（Run の前処理や自動 Push など、別途結果を出す場合）。 */
  silent?: boolean;
  /** Office ファイルの保存を強制/抑制する（未指定なら設定値 saveAfterPush に従う）。 */
  save?: boolean;
  /** 進捗表示の出し方。既定は "notification"。"none" は無表示（自動 Push 用）。 */
  progress?: PushProgressKind;
  /** 対象が開いているインスタンスにのみ接続し、閉じていれば起動せずスキップする。 */
  attachOnly?: boolean;
  /** 成功時（スキップ含む）にブリッジ結果を受け取る。 */
  onResult?: (data: { pushed: number; skipped?: string }) => void;
  /** 失敗時の処理を差し替える（未指定なら既定のエラーダイアログを表示）。 */
  onError?: (err: unknown) => void;
}

/**
 * 指定レコード群を Push する共通処理。
 * @returns 成功（またはスキップ）なら true、失敗（ブリッジエラー）なら false。
 */
export async function runPush(
  dir: string,
  manifest: Manifest,
  records: ComponentRecord[],
  opts?: RunPushOptions
): Promise<boolean> {
  const cfg = getConfig();
  const fileName = path.basename(manifest.source);
  const progressKind: PushProgressKind = opts?.progress ?? "notification";

  // VSCode 側でファイル名(=モジュール名)を二重化/typo したまま Push すると、Office 側で
  // 同名モジュールが重複してコンパイル不能になり Push/Pull が反映されない。事前に検出して止める。
  const dups = findDuplicateComponentNames(records);
  if (dups.length > 0) {
    const err = new Error(duplicateNamesMessage(dups));
    if (opts?.onError) opts.onError(err);
    else await handleBridgeError(err, manifest.app);
    return false;
  }

  // ファイル読み込み → ブリッジ呼び出し本体（進捗表示の有無で包み方を変える）。
  const work = async (
    token?: vscode.CancellationToken
  ): Promise<{
    hashes: Map<string, string>;
    data: { pushed: number; skipped?: string };
  }> => {
    const components: PushComponent[] = [];
    const newHashes = new Map<string, string>();

    for (const rec of records) {
      const filePath = path.join(dir, rec.file);
      let text: string;
      try {
        text = await fs.readFile(filePath, "utf8");
      } catch {
        continue; // ファイルが無ければスキップ
      }
      newHashes.set(rec.name, sha256(text));
      components.push(await buildPushComponent(dir, rec, text));
    }

    if (components.length === 0) {
      throw new Error("書き戻す対象ファイルが見つかりませんでした。");
    }

    const { data } = await invoke<{ pushed: number; skipped?: string }>(
      scriptFor(manifest.app, "push"),
      {
        path: manifest.source,
        headless: cfg.headless,
        save: opts?.save ?? cfg.saveAfterPush,
        attachOnly: opts?.attachOnly ?? false,
        components,
      } as PushPayload,
      bridgeOptions(cfg),
      token
    );
    return { hashes: newHashes, data };
  };

  try {
    let result: {
      hashes: Map<string, string>;
      data: { pushed: number; skipped?: string };
    };
    if (progressKind === "none") {
      result = await work();
    } else {
      result = await vscode.window.withProgress(
        {
          location:
            progressKind === "window"
              ? vscode.ProgressLocation.Window
              : vscode.ProgressLocation.Notification,
          title: `Perfect VBA: ${fileName} へ書き戻し中...`,
          cancellable: progressKind === "notification",
        },
        (_progress, token) => work(token)
      );
    }

    // attachOnly で対象が未起動だった場合は、マニフェストを更新せず結果のみ通知する。
    if (result.data.skipped === "NOT_OPEN") {
      opts?.onResult?.(result.data);
      return true;
    }

    const updated = result.hashes;

    // VSCode 側で新規作成し今回 Push したモジュールを manifest に登録する。
    for (const rec of records) {
      if (!manifest.components.some((c) => c.name === rec.name)) {
        manifest.components.push(rec);
      }
    }
    // manifest のハッシュを更新
    for (const rec of manifest.components) {
      const h = updated.get(rec.name);
      if (h) rec.hash = h;
    }
    await writeManifest(dir, manifest);
    await vscode.commands.executeCommand("perfectVba.refreshTree");

    opts?.onResult?.(result.data);

    if (!opts?.silent) {
      vscode.window.showInformationMessage(
        `Perfect VBA: ${fileName} へ ${updated.size} 個のコンポーネントを書き戻しました。`
      );
    }
    return true;
  } catch (err) {
    if (opts?.onError) opts.onError(err);
    else await handleBridgeError(err, manifest.app);
    return false;
  }
}

async function buildPushComponent(
  dir: string,
  rec: ComponentRecord,
  text: string
): Promise<PushComponent> {
  const kind = rec.kind as ComponentKind;
  if (kind === "document") {
    return { name: rec.name, kind, codeText: text };
  }
  const buf = encodeForWrite(text, rec.enc);
  const pc: PushComponent = {
    name: rec.name,
    kind,
    contentB64: buf.toString("base64"),
    fileExt: importExt(kind),
  };
  // std/class は「既存モジュールへ CodeModule 直接置換」で反映する（Remove+Import の
  // 遅延・重複生成や、VBE でコードウィンドウを開いたまま Import すると no-op 化する
  // 問題を避ける）。新規モジュール作成時のみ contentB64 を Import する。
  if (kind === "std" || kind === "class") {
    pc.codeText = text;
  }
  if (kind === "form" && rec.frxFile) {
    try {
      const frx = await fs.readFile(path.join(dir, rec.frxFile));
      pc.frxB64 = frx.toString("base64");
    } catch {
      /* .frx が無ければそのまま（コードのみ） */
    }
  }
  return pc;
}
