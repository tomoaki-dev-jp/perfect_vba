// Push 前のモジュール名の検証。
// VSCode 側でファイル名(=モジュール名)を二重化・typo すると、Office 側で同名モジュールが
// 重複してコンパイル不能（「あいまいな名前が検出されました」「プロシージャが見つかりません」）
// になり、Push/Pull が正しく反映されない。これを Push 前に検出して明確に止める。
import { ComponentRecord } from "./manifest";

/**
 * VBA モジュール名を比較用に正規化する。
 * VBA の識別子は大文字小文字を区別しないため小文字化し、前後空白とユニコード合成の揺れ
 * （濁点の合成/分解など）を吸収する。全角/半角は VBA 上で別名のため正規化しない（NFKC は使わない）。
 */
export function normalizeModuleName(name: string): string {
  return name.normalize("NFC").trim().toLowerCase();
}

export interface DuplicateGroup {
  /** 正規化後の名前。 */
  key: string;
  /** 同じ正規化名を持つレコード群（2 件以上）。 */
  records: ComponentRecord[];
}

/**
 * Push 対象レコードから、正規化後に名前が衝突するグループを返す。
 * 同名・別ファイル（例: Foo.bas と Foo.cls、Foo.bas を複製した "Foo - コピー" を Foo にした等）を検出する。
 */
export function findDuplicateComponentNames(
  records: ComponentRecord[]
): DuplicateGroup[] {
  const map = new Map<string, ComponentRecord[]>();
  for (const rec of records) {
    const key = normalizeModuleName(rec.name);
    const arr = map.get(key);
    if (arr) arr.push(rec);
    else map.set(key, [rec]);
  }
  const groups: DuplicateGroup[] = [];
  for (const [key, recs] of map) {
    if (recs.length > 1) groups.push({ key, records: recs });
  }
  return groups;
}

/** 重複検出結果を、ユーザー向けの日本語メッセージにする。 */
export function duplicateNamesMessage(groups: DuplicateGroup[]): string {
  const detail = groups
    .map((g) => `"${g.records[0].name}"（${g.records.map((r) => r.file).join(", ")}）`)
    .join("; ");
  return (
    `モジュール名が重複しています: ${detail}。` +
    `同じ名前のモジュールが複数あると Access/Excel 側でコンパイルできず` +
    `（「あいまいな名前が検出されました」「プロシージャが見つかりません」）、` +
    `Push/Pull が正しく反映されません。重複したファイルを 1 つに整理し、` +
    `ファイル名を正しいモジュール名に直してから再実行してください。`
  );
}
