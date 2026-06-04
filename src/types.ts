// 拡張機能と PowerShell ブリッジ間でやり取りする共通の型定義。

export type AppKind = "excel" | "access";

/**
 * VBA コンポーネントの種別。
 * - std/class/form: VBE の Export/Import で扱う（.bas/.cls/.frm[+.frx]）
 * - document: ThisWorkbook / Sheet / フォーム・レポートのコードビハインド。
 *   Remove/Import 不可のため CodeModule の行置換で扱う。
 * - accForm/accReport/accMacro: Access の SaveAsText/LoadFromText で扱う（定義＋コード全体）。
 */
export type ComponentKind =
  | "std"
  | "class"
  | "form"
  | "document"
  | "accForm"
  | "accReport"
  | "accMacro";

/** バイト列ベースで保存する（＝エンコーディング変換が必要な）種別か。 */
export function isByteKind(kind: ComponentKind): boolean {
  return kind !== "document";
}

/** ブリッジが返すエラー。code はプログラムで分岐するための安定識別子。 */
export interface BridgeError {
  code: string;
  message: string;
}

/** PowerShell スクリプトが stdout に出力する単一 JSON の形。 */
export interface BridgeResult<T = unknown> {
  ok: boolean;
  data?: T;
  /** システムの ANSI コードページ番号（例: 932）。auto エンコーディング解決に使う。 */
  acp?: number;
  error?: BridgeError;
}

/** list 操作の戻り（コンポーネント一覧のみ）。 */
export interface ListResult {
  app: AppKind;
  projectName: string;
  components: Array<{
    name: string;
    kind: ComponentKind;
    hasFrx?: boolean;
  }>;
}

/** pull 操作で 1 コンポーネントが返す内容。 */
export interface PulledComponent {
  name: string;
  kind: ComponentKind;
  /** byte 種別の生バイト（base64）。エンコーディング未変換のまま。 */
  contentB64?: string;
  /** form の .frx バイナリ（base64）。 */
  frxB64?: string;
  /** document 種別のコード（Unicode 文字列そのまま）。 */
  codeText?: string;
}

export interface PullResult {
  app: AppKind;
  projectName: string;
  components: PulledComponent[];
}

/** push 操作で 1 コンポーネントを送る内容。 */
export interface PushComponent {
  name: string;
  kind: ComponentKind;
  /** byte 種別: 対象エンコーディングへ変換済みのバイト（base64）。 */
  contentB64?: string;
  /** form の .frx バイナリ（base64、無変更でそのまま戻す）。 */
  frxB64?: string;
  /**
   * 反映するコード（Unicode 文字列）。
   * - document: そのまま行置換する本体。
   * - std/class: エクスポート形式の全文。既存モジュールへ CodeModule 直接置換する際に使う
   *   （ブリッジ側でヘッダを除去）。新規作成時は contentB64 の Import が使われる。
   */
  codeText?: string;
  /** Import 用一時ファイルの拡張子（".bas"/".cls"/".frm"/".txt"）。 */
  fileExt?: string;
}

export interface PullPayload {
  path: string;
  headless: boolean;
  includeDocumentModules: boolean;
}

export interface PushPayload {
  path: string;
  headless: boolean;
  save: boolean;
  components: PushComponent[];
}

/** run 操作で渡す内容（実行するプロシージャの指定）。 */
export interface RunPayload {
  path: string;
  headless: boolean;
  /** 実行するプロシージャ名。 */
  macro: string;
  /** Excel: モジュール修飾用の標準モジュール名。Access では未使用。 */
  module?: string;
}

/** run 操作の戻り。 */
export interface RunResult {
  ran: boolean;
  /** 実行したマクロ参照（"Module.Proc" 等）。 */
  macro: string;
  /** Function の戻り値を文字列化したもの（Sub・戻り値なしは null）。 */
  returnValue?: string | null;
}
