// Pull: Office ファイル → ワークスペースの VBA テキストファイル群 + manifest。
import * as vscode from "vscode";
import * as path from "path";
import * as fs from "fs/promises";
import { getConfig, bridgeOptions } from "../config";
import { invoke } from "../bridge/powershell";
import { appForFile, scriptFor } from "../bridge/office";
import { PullPayload, PullResult } from "../types";
import {
  resolveCodepage,
  decodeBytes,
  sniffEncoding,
  EncodingSetting,
} from "../encoding";
import {
  ComponentRecord,
  Manifest,
  MANIFEST_VERSION,
  projectFolderName,
  readManifest,
  sha256,
  writeManifest,
} from "../model/manifest";
import { componentFileName } from "../model/layout";
import {
  handleBridgeError,
  pickOfficeFile,
  pickProject,
  projectsRootDir,
  safeUnlink,
} from "./shared";

interface RefreshArg {
  projectDir?: string;
  manifestSource?: string;
}

/** ツリー項目やコマンドパレットからの再取り込み（再 Pull）。 */
export async function refreshCommand(arg?: unknown): Promise<void> {
  let source: string | undefined;
  if (arg && typeof arg === "object") {
    const a = arg as RefreshArg;
    source = a.manifestSource;
    if (!source && a.projectDir) {
      const m = await readManifest(a.projectDir);
      source = m?.source;
    }
  }
  if (!source) {
    const entry = await pickProject();
    source = entry?.manifest.source;
  }
  if (!source) return;
  await pullCommand(source);
}

export async function pullCommand(arg?: vscode.Uri | string): Promise<void> {
  let source = typeof arg === "string" ? arg : arg?.fsPath;
  if (!source) source = await pickOfficeFile();
  if (!source) return;

  const app = appForFile(source);
  if (!app) {
    vscode.window.showErrorMessage(`Perfect VBA: 対応していないファイル形式です: ${source}`);
    return;
  }
  if (!projectsRootDir()) {
    vscode.window.showErrorMessage("Perfect VBA: 先にフォルダ（ワークスペース）を開いてください。");
    return;
  }

  const cfg = getConfig();
  const fileName = path.basename(source);
  try {
    await vscode.window.withProgress(
      {
        location: vscode.ProgressLocation.Notification,
        title: `Perfect VBA: ${fileName} を取り込み中...`,
        cancellable: true,
      },
      async (_progress, token) => {
        const { data, acp } = await invoke<PullResult>(
          scriptFor(app, "pull"),
          {
            path: source,
            headless: cfg.headless,
            includeDocumentModules: cfg.includeDocumentModules,
          } as PullPayload,
          bridgeOptions(cfg),
          token
        );
        await writePulled(source!, data, acp, cfg.encoding);
      }
    );
    vscode.window.showInformationMessage(
      `Perfect VBA: ${fileName} を取り込みました。`
    );
    await vscode.commands.executeCommand("perfectVba.refreshTree");
  } catch (err) {
    await handleBridgeError(err, app);
  }
}

async function writePulled(
  source: string,
  data: PullResult,
  acp: number | undefined,
  encodingSetting: EncodingSetting
): Promise<void> {
  const root = projectsRootDir()!;
  const dir = path.join(root, projectFolderName(source));
  await fs.mkdir(dir, { recursive: true });

  const prev = await readManifest(dir);
  const cp = resolveCodepage(encodingSetting, acp);

  const records: ComponentRecord[] = [];
  const written = new Set<string>();

  for (const comp of data.components) {
    const { file, frxFile } = componentFileName(comp.name, comp.kind);
    let text: string;
    let enc: string;
    let recFrx: string | undefined;

    if (comp.kind === "document") {
      text = comp.codeText ?? "";
      enc = "document";
    } else {
      const buf = Buffer.from(comp.contentB64 ?? "", "base64");
      enc = sniffEncoding(buf, cp);
      text = decodeBytes(buf, enc);
      if (comp.kind === "form" && frxFile && comp.frxB64) {
        await fs.writeFile(path.join(dir, frxFile), Buffer.from(comp.frxB64, "base64"));
        written.add(frxFile);
        recFrx = frxFile;
      }
    }

    await fs.writeFile(path.join(dir, file), text, "utf8");
    written.add(file);
    records.push({
      name: comp.name,
      kind: comp.kind,
      file,
      frxFile: recFrx,
      enc,
      hash: sha256(text),
    });
  }

  // 旧 manifest にあって今回現れなかったファイルを削除（リネーム/削除の追従）
  if (prev) {
    for (const old of prev.components) {
      if (!written.has(old.file)) await safeUnlink(path.join(dir, old.file));
      if (old.frxFile && !written.has(old.frxFile)) {
        await safeUnlink(path.join(dir, old.frxFile));
      }
    }
  }

  const manifest: Manifest = {
    version: MANIFEST_VERSION,
    app: data.app,
    source,
    projectName: data.projectName,
    encodingSetting,
    pulledAt: new Date().toISOString(),
    components: records,
  };
  await writeManifest(dir, manifest);
}
