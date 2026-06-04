// プロジェクトフォルダごとの .perfect-vba.json（コンポーネント名⇔ファイル⇔種別・エンコーディング・ハッシュ）。
import * as fs from "fs/promises";
import * as path from "path";
import { createHash } from "crypto";
import { AppKind, ComponentKind } from "../types";

export const MANIFEST_NAME = ".perfect-vba.json";
export const MANIFEST_VERSION = 1;

export interface ComponentRecord {
  name: string;
  kind: ComponentKind;
  /** プロジェクトフォルダ基準の相対ファイル名（例: "Module1.bas"）。 */
  file: string;
  /** form のみ: 対の .frx ファイル名。 */
  frxFile?: string;
  /** byte 種別の保存エンコーディング名（iconv-lite 名）。document は "document"。 */
  enc: string;
  /** pull/push 時点の UTF-8 ワークスペース内容の sha256（外部編集・競合検出用）。 */
  hash: string;
}

export interface Manifest {
  version: number;
  app: AppKind;
  /** 取り込み元 Office ファイルの絶対パス。 */
  source: string;
  projectName: string;
  /** 取り込み時の encoding 設定値。 */
  encodingSetting: string;
  pulledAt: string;
  components: ComponentRecord[];
}

export function sha256(text: string): string {
  return createHash("sha256").update(text, "utf8").digest("hex");
}

export async function readManifest(dir: string): Promise<Manifest | undefined> {
  const file = path.join(dir, MANIFEST_NAME);
  try {
    const raw = await fs.readFile(file, "utf8");
    return JSON.parse(raw) as Manifest;
  } catch {
    return undefined;
  }
}

export async function writeManifest(dir: string, manifest: Manifest): Promise<void> {
  await fs.mkdir(dir, { recursive: true });
  const file = path.join(dir, MANIFEST_NAME);
  await fs.writeFile(file, JSON.stringify(manifest, null, 2), "utf8");
}

/** Office ファイルのパスから、プロジェクトフォルダ名（サニタイズ済み）を作る。 */
export function projectFolderName(sourcePath: string): string {
  return path.basename(sourcePath).replace(/[<>:"/\\|?*]/g, "_");
}
