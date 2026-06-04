// VBA ソースから手続き(Sub/Function)宣言を抽出し、実行可能なマクロを絞り込む。
// Application.Run で名前実行できるのは「Private でない・必須引数を持たない Sub/Function」
// （標準モジュールにあるもの = Excel の Alt+F8 マクロ一覧に出るもの）。

export type ProcKind = "Sub" | "Function";
export type ProcVisibility = "Public" | "Private" | "Friend" | "Default";

export interface VbaProcedure {
  name: string;
  kind: ProcKind;
  visibility: ProcVisibility;
  /** Optional / ParamArray でない引数（=呼び出し時に値が必須）を持つか。 */
  hasRequiredParams: boolean;
  /** 宣言が始まる 0 始まりの行番号。 */
  line: number;
}

// 行頭の任意の可視性 + 任意の Static + Sub/Function + 名前 + 任意の引数リスト。
// "Declare Sub ..." は可視性/Static の直後が Sub/Function でないため一致しない（=除外）。
// "Property Get/Let/Set" は種別が Sub/Function でないため一致しない（=除外）。
const DECL =
  /^[ \t]*(?:(Public|Private|Friend|Global)[ \t]+)?(?:Static[ \t]+)?(Sub|Function)[ \t]+([A-Za-z_][A-Za-z0-9_]*)[ \t]*(?:\(([^)]*)\))?/i;

/** 行継続(" _")を畳んだ論理行と、その開始行番号(0 始まり)を返す。 */
function logicalLines(source: string): { text: string; line: number }[] {
  const raw = source.split(/\r\n|\r|\n/);
  const out: { text: string; line: number }[] = [];
  let i = 0;
  while (i < raw.length) {
    let line = raw[i];
    const start = i;
    while (/\s_[ \t]*$/.test(line) && i + 1 < raw.length) {
      line = line.replace(/\s_[ \t]*$/, " ") + raw[i + 1];
      i++;
    }
    out.push({ text: line, line: start });
    i++;
  }
  return out;
}

function classifyVisibility(raw: string | undefined): ProcVisibility {
  switch ((raw ?? "").toLowerCase()) {
    case "private":
      return "Private";
    case "friend":
      return "Friend";
    case "public":
    case "global":
      return "Public";
    default:
      return "Default";
  }
}

function hasRequiredParams(params: string): boolean {
  const trimmed = params.trim();
  if (!trimmed) return false;
  // 引数の区切りは（配列の () 内にカンマが来ないため）単純なカンマ分割で十分。
  return trimmed.split(",").some((p) => {
    const t = p.trim();
    if (!t) return false;
    if (/^Optional\b/i.test(t)) return false;
    if (/^ParamArray\b/i.test(t)) return false;
    return true;
  });
}

export function parseProcedures(source: string): VbaProcedure[] {
  const procs: VbaProcedure[] = [];
  for (const { text, line } of logicalLines(source)) {
    if (/^[ \t]*'/.test(text)) continue; // コメント行
    if (/^[ \t]*Rem\b/i.test(text)) continue; // Rem コメント
    const m = DECL.exec(text);
    if (!m) continue;
    procs.push({
      name: m[3],
      kind: m[2].toLowerCase() === "function" ? "Function" : "Sub",
      visibility: classifyVisibility(m[1]),
      hasRequiredParams: hasRequiredParams(m[4] ?? ""),
      line,
    });
  }
  return procs;
}

/** Application.Run で名前実行できる候補（Private でない・必須引数なし）を返す。 */
export function runnableMacros(procs: VbaProcedure[]): VbaProcedure[] {
  return procs.filter((p) => p.visibility !== "Private" && !p.hasRequiredParams);
}
