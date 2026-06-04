// PowerShell COM ブリッジ。.ps1 を子プロセス起動し、payload を stdin(JSON)で渡し、
// stdout から単一 JSON(BridgeResult)を受け取る。エラーは型付き例外に写像する。
import { spawn } from "child_process";
import * as path from "path";
import * as vscode from "vscode";
import { BridgeResult } from "../types";

/** スクリプトは dist/extension.js と同じ階層の scripts/ に配置される。 */
const SCRIPTS_DIR = path.join(__dirname, "scripts");

export class BridgeError extends Error {
  constructor(public code: string, message: string) {
    super(message);
    this.name = "BridgeError";
  }
}

export interface BridgeOptions {
  powerShellPath: string;
}

/**
 * .ps1 を実行し BridgeResult を返す（ok=false でも resolve する。プロセス異常時のみ reject）。
 */
export function runScript<T>(
  scriptName: string,
  payload: unknown,
  opts: BridgeOptions,
  token?: vscode.CancellationToken
): Promise<BridgeResult<T>> {
  const scriptPath = path.join(SCRIPTS_DIR, scriptName);
  const args = [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
  ];

  return new Promise((resolve, reject) => {
    const child = spawn(opts.powerShellPath || "powershell.exe", args, {
      windowsHide: true,
    });
    let stdout = "";
    let stderr = "";

    const cancelSub = token?.onCancellationRequested(() => {
      child.kill();
      reject(new BridgeError("CANCELLED", "操作がキャンセルされました。"));
    });

    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (d) => (stdout += d));
    child.stderr.on("data", (d) => (stderr += d));

    child.on("error", (err) => {
      cancelSub?.dispose();
      reject(
        new BridgeError(
          "SPAWN_FAILED",
          `PowerShell の起動に失敗しました (${opts.powerShellPath}): ${err.message}`
        )
      );
    });

    child.on("close", (exitCode) => {
      cancelSub?.dispose();
      const trimmed = stdout.trim();
      if (!trimmed) {
        reject(
          new BridgeError(
            "NO_OUTPUT",
            `スクリプトが出力を返しませんでした (exit=${exitCode}).\nstderr: ${stderr.trim()}`
          )
        );
        return;
      }
      try {
        const parsed = JSON.parse(extractJson(trimmed)) as BridgeResult<T>;
        resolve(parsed);
      } catch (e) {
        reject(
          new BridgeError(
            "BAD_OUTPUT",
            `スクリプト出力の JSON 解析に失敗しました: ${(e as Error).message}\n` +
              `--- stdout ---\n${trimmed}\n--- stderr ---\n${stderr.trim()}`
          )
        );
      }
    });

    child.stdin.write(JSON.stringify(payload ?? {}));
    child.stdin.end();
  });
}

/** runScript を実行し、ok=false なら BridgeError を投げて data を返す。 */
export async function invoke<T>(
  scriptName: string,
  payload: unknown,
  opts: BridgeOptions,
  token?: vscode.CancellationToken
): Promise<{ data: T; acp?: number }> {
  const res = await runScript<T>(scriptName, payload, opts, token);
  if (!res.ok || res.data === undefined) {
    const code = res.error?.code ?? "UNKNOWN";
    const message = res.error?.message ?? "不明なエラーが発生しました。";
    throw new BridgeError(code, message);
  }
  return { data: res.data, acp: res.acp };
}

/** stdout 末尾の JSON オブジェクトを抽出（前後のノイズ行に対する保険）。 */
function extractJson(s: string): string {
  const start = s.indexOf("{");
  const end = s.lastIndexOf("}");
  if (start >= 0 && end >= start) {
    return s.slice(start, end + 1);
  }
  return s;
}
