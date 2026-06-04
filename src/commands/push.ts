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

/**
 * 指定レコード群を Push する共通処理。
 * @param opts.silent 成功トーストを抑制する（Run の前処理など、別途結果を出す場合）。
 * @param opts.save   Office ファイルの保存を強制/抑制する（未指定なら設定値に従う）。
 * @returns 成功したら true、失敗（ブリッジエラー）なら false。
 */
export async function runPush(
  dir: string,
  manifest: Manifest,
  records: ComponentRecord[],
  opts?: { silent?: boolean; save?: boolean }
): Promise<boolean> {
  const cfg = getConfig();
  const fileName = path.basename(manifest.source);

  try {
    const updated = await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: `Perfect VBA: ${fileName} へ書き戻し中...`,
        cancellable: true,
      },
      async (_progress, token) => {
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

        await invoke<{ pushed: number }>(
          scriptFor(manifest.app, "push"),
          {
            path: manifest.source,
            headless: cfg.headless,
            save: opts?.save ?? cfg.saveAfterPush,
            components,
          } as PushPayload,
          bridgeOptions(cfg),
          token
        );
        return newHashes;
      }
    );

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

    if (!opts?.silent) {
      vscode.window.showInformationMessage(
        `Perfect VBA: ${fileName} へ ${updated.size} 個のコンポーネントを書き戻しました。`
      );
    }
    return true;
  } catch (err) {
    await handleBridgeError(err, manifest.app);
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
