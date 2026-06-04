// コンポーネント種別 → ワークスペース上のファイル名 / Import 用拡張子 の対応。
import { ComponentKind } from "../types";

/** 種別と名前から、保存ファイル名（と form の .frx 名）を決める。 */
export function componentFileName(
  name: string,
  kind: ComponentKind
): { file: string; frxFile?: string } {
  const safe = sanitizeName(name);
  switch (kind) {
    case "std":
      return { file: `${safe}.bas` };
    case "class":
      return { file: `${safe}.cls` };
    case "document":
      return { file: `${safe}.cls` };
    case "form":
      return { file: `${safe}.frm`, frxFile: `${safe}.frx` };
    case "accForm":
      return { file: `Form_${safe}.txt` };
    case "accReport":
      return { file: `Report_${safe}.txt` };
    case "accMacro":
      return { file: `Macro_${safe}.txt` };
  }
}

/** push 時に PowerShell へ渡す Import 用一時ファイルの拡張子。 */
export function importExt(kind: ComponentKind): string {
  switch (kind) {
    case "std":
      return ".bas";
    case "class":
      return ".cls";
    case "form":
      return ".frm";
    case "accForm":
    case "accReport":
    case "accMacro":
      return ".txt";
    case "document":
      return ".cls"; // 実際には Import せず行置換
  }
}

export function sanitizeName(name: string): string {
  return name.replace(/[<>:"/\\|?*]/g, "_");
}
