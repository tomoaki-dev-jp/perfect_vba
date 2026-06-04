// ファイル拡張子 → アプリ種別、操作 → スクリプト名 の対応。
import * as path from "path";
import { AppKind } from "../types";

const EXCEL_EXT = new Set([".xlsm", ".xlsb", ".xlam", ".xls", ".xltm", ".xlt"]);
const ACCESS_EXT = new Set([".accdb", ".mdb", ".accda", ".accde", ".mde", ".accdr"]);

export function appForFile(file: string): AppKind | undefined {
  const ext = path.extname(file).toLowerCase();
  if (EXCEL_EXT.has(ext)) return "excel";
  if (ACCESS_EXT.has(ext)) return "access";
  return undefined;
}

export function scriptFor(
  app: AppKind,
  op: "list" | "pull" | "push" | "run"
): string {
  return `${app}-${op}.ps1`;
}

/** ファイル選択ダイアログ用のフィルタ。 */
export const OFFICE_FILTERS: { [name: string]: string[] } = {
  "Office VBA ファイル": [
    "xlsm",
    "xlsb",
    "xlam",
    "xls",
    "xltm",
    "accdb",
    "mdb",
    "accda",
    "accde",
    "mde",
  ],
  Excel: ["xlsm", "xlsb", "xlam", "xls", "xltm"],
  Access: ["accdb", "mdb", "accda", "accde", "mde"],
};
