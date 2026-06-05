// 保存時の自動 Push（リアルタイム反映）。
// VBA ファイルを保存すると、開いている Office インスタンスへサイレントに書き戻す。
// - attachOnly: 対象が開かれていなければ Office を起動せずスキップする。
// - 既定では Office ファイルを保存しない（起動中の VBProject にコードだけ即反映）。
// - 通知は出さず、状態はステータスバーに集約する。
import * as vscode from "vscode";
import * as path from "path";
import { getConfig } from "./config";
import { newRecordForFile, runPush } from "./commands/push";
import {
  findProjectContainingPath,
  findProjectForFile,
  handleBridgeError,
} from "./commands/shared";
import { BridgeError } from "./bridge/powershell";
import { ComponentRecord, Manifest } from "./model/manifest";
import { ComponentKind } from "./types";

/** 保存バーストをまとめる待ち時間（ms）。 */
const DEBOUNCE_MS = 150;
/** 同期/未起動表示をアイドルへ戻すまでの時間（ms）。 */
const RESET_MS = 4000;
/** 自動 Push の対象になりうる拡張子（最終判定はプロジェクト所属・種別で行う）。 */
const PUSH_EXTS = new Set([".bas", ".cls"]);
/**
 * 自動 Push の対象とする種別。CodeModule 直接置換で即時・安全に反映できるものに限定する。
 * フォーム(Remove+Import)や Access のフォーム/レポート/マクロ(LoadFromText)は
 * 「オブジェクト全体の上書き」になり、保存のたびに自動実行するのは危険なため除外する
 * （これらは従来どおり手動 Push で書き戻す）。
 */
const AUTO_PUSH_KINDS = new Set<ComponentKind>(["std", "class", "document"]);

interface ProjectBatch {
  manifest: Manifest;
  recs: Map<string, ComponentRecord>;
}

/**
 * 保存イベントを購読し、対象プロジェクトごとにまとめて自動 Push する。
 * Push 同士は直列化し、COM 自動化の競合を避ける。
 */
export class AutoPushController implements vscode.Disposable {
  private readonly status: vscode.StatusBarItem;
  private readonly disposables: vscode.Disposable[] = [];
  private pending = new Map<string, ProjectBatch>();
  private debounce?: ReturnType<typeof setTimeout>;
  private resetTimer?: ReturnType<typeof setTimeout>;
  private running = false;
  private trustPrompted = false;

  constructor() {
    this.status = vscode.window.createStatusBarItem(
      vscode.StatusBarAlignment.Right,
      99
    );
    this.status.command = "perfectVba.toggleAutoPush";
    this.disposables.push(
      this.status,
      vscode.workspace.onDidSaveTextDocument((doc) => this.onSave(doc)),
      vscode.workspace.onDidChangeConfiguration((e) => {
        if (e.affectsConfiguration("perfectVba.autoPushOnSave")) {
          this.trustPrompted = false;
          this.renderIdle();
        }
      })
    );
    this.renderIdle();
  }

  private enabled(): boolean {
    return getConfig().autoPushOnSave;
  }

  private onSave(doc: vscode.TextDocument): void {
    if (!this.enabled()) return;
    if (doc.uri.scheme !== "file") return;
    if (!PUSH_EXTS.has(path.extname(doc.uri.fsPath).toLowerCase())) return;
    void this.enqueue(doc.uri.fsPath);
  }

  /** 保存されたファイルを所属プロジェクト・コンポーネントへ解決し、バッチに積む。 */
  private async enqueue(file: string): Promise<void> {
    const resolved = await this.resolve(file);
    if (!resolved) return; // 取り込み済みプロジェクト外のファイルは無視
    const { dir, manifest, rec } = resolved;
    let batch = this.pending.get(dir);
    if (!batch) {
      batch = { manifest, recs: new Map() };
      this.pending.set(dir, batch);
    }
    batch.manifest = manifest;
    batch.recs.set(rec.name, rec);
    if (this.debounce) clearTimeout(this.debounce);
    this.debounce = setTimeout(() => void this.flush(), DEBOUNCE_MS);
  }

  private async resolve(
    file: string
  ): Promise<
    { dir: string; manifest: Manifest; rec: ComponentRecord } | undefined
  > {
    const found = await findProjectForFile(file);
    if (found) {
      const rec = found.entry.manifest.components.find(
        (c) => c.name === found.componentName
      );
      if (rec && AUTO_PUSH_KINDS.has(rec.kind)) {
        return { dir: found.entry.dir, manifest: found.entry.manifest, rec };
      }
      return undefined;
    }
    // manifest 未登録でも、プロジェクト配下の .bas/.cls なら新規モジュールとして扱う。
    const proj = await findProjectContainingPath(file);
    if (proj) {
      const rec = newRecordForFile(proj.manifest, path.basename(file));
      if (rec && AUTO_PUSH_KINDS.has(rec.kind)) {
        return { dir: proj.dir, manifest: proj.manifest, rec };
      }
    }
    return undefined;
  }

  /** たまったバッチをプロジェクト単位で直列に Push する。 */
  private async flush(): Promise<void> {
    if (this.running) {
      // 実行中に届いた分は、現在の Push 完了後に処理し直す。
      this.debounce = setTimeout(() => void this.flush(), DEBOUNCE_MS);
      return;
    }
    if (this.pending.size === 0) return;
    this.running = true;
    const batch = this.pending;
    this.pending = new Map();
    const cfg = getConfig();
    try {
      for (const [dir, b] of batch) {
        const records = [...b.recs.values()];
        this.renderPushing(records.length);
        await runPush(dir, b.manifest, records, {
          silent: true,
          save: cfg.autoPushSaveOffice,
          attachOnly: true,
          progress: "none",
          onResult: (data) => this.renderResult(data, records.length),
          onError: (err) => this.renderError(err, b.manifest.app),
        });
      }
    } finally {
      this.running = false;
      if (this.pending.size > 0) {
        this.debounce = setTimeout(() => void this.flush(), 50);
      }
    }
  }

  // ---- ステータスバー表示 ----

  private renderIdle(): void {
    this.clearReset();
    if (!this.enabled()) {
      this.status.hide();
      return;
    }
    this.status.text = "$(sync) VBA 自動Push";
    this.status.tooltip =
      "保存時に、開いている Office へ自動 Push します（クリックで無効化）";
    this.status.backgroundColor = undefined;
    this.status.show();
  }

  private renderPushing(n: number): void {
    this.clearReset();
    this.status.text = `$(sync~spin) VBA Push中… (${n})`;
    this.status.tooltip = "開いている Office へ書き戻し中…";
    this.status.backgroundColor = undefined;
    this.status.show();
  }

  private renderResult(
    data: { pushed: number; skipped?: string },
    n: number
  ): void {
    if (data.skipped === "NOT_OPEN") {
      this.status.text = "$(circle-slash) VBA 未起動";
      this.status.tooltip =
        "対象の Office ファイルが開かれていないため自動 Push をスキップしました。手動 Push なら起動して書き戻せます。";
    } else {
      const t = new Date().toLocaleTimeString();
      this.status.text = `$(check) VBA 同期 ${t}`;
      this.status.tooltip = `${n} 個のコンポーネントを、開いている Office へ反映しました（${t}）`;
    }
    this.status.backgroundColor = undefined;
    this.status.show();
    this.scheduleIdle(RESET_MS);
  }

  private renderError(err: unknown, app: Manifest["app"]): void {
    const msg = err instanceof Error ? err.message : String(err);
    this.status.text = "$(error) VBA Push失敗";
    this.status.tooltip = `自動 Push に失敗しました: ${msg}`;
    this.status.backgroundColor = new vscode.ThemeColor(
      "statusBarItem.errorBackground"
    );
    this.status.show();
    this.scheduleIdle(8000);

    // Trust 未許可は対処が必要なので一度だけ案内する（毎保存ごとのダイアログを抑止）。
    if (
      err instanceof BridgeError &&
      err.code === "TRUST_DISABLED" &&
      !this.trustPrompted
    ) {
      this.trustPrompted = true;
      void handleBridgeError(err, app);
    }
  }

  private scheduleIdle(ms: number): void {
    this.clearReset();
    this.resetTimer = setTimeout(() => this.renderIdle(), ms);
  }

  private clearReset(): void {
    if (this.resetTimer) {
      clearTimeout(this.resetTimer);
      this.resetTimer = undefined;
    }
  }

  dispose(): void {
    if (this.debounce) clearTimeout(this.debounce);
    this.clearReset();
    for (const d of this.disposables) d.dispose();
  }
}

/** 設定 `perfectVba.autoPushOnSave` をトグルする。 */
export async function toggleAutoPushCommand(): Promise<void> {
  const cfg = vscode.workspace.getConfiguration("perfectVba");
  const cur = cfg.get<boolean>("autoPushOnSave", false);
  const target =
    vscode.workspace.workspaceFolders &&
    vscode.workspace.workspaceFolders.length > 0
      ? vscode.ConfigurationTarget.Workspace
      : vscode.ConfigurationTarget.Global;
  await cfg.update("autoPushOnSave", !cur, target);
  vscode.window.showInformationMessage(
    `Perfect VBA: 保存時の自動 Push（リアルタイム反映）を${
      !cur ? "有効" : "無効"
    }にしました。`
  );
}
