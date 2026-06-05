import * as vscode from "vscode";
import { EncodingSetting } from "./encoding";
import { BridgeOptions } from "./bridge/powershell";

export interface PerfectVbaConfig {
  workspaceRoot: string;
  encoding: EncodingSetting;
  headless: boolean;
  powerShellPath: string;
  saveAfterPush: boolean;
  includeDocumentModules: boolean;
  pushBeforeRun: boolean;
  autoPushOnSave: boolean;
  autoPushSaveOffice: boolean;
}

export function getConfig(): PerfectVbaConfig {
  const c = vscode.workspace.getConfiguration("perfectVba");
  return {
    workspaceRoot: c.get<string>("workspaceRoot", ".vba"),
    encoding: c.get<EncodingSetting>("encoding", "auto"),
    headless: c.get<boolean>("headless", false),
    powerShellPath: c.get<string>("powerShellPath", "powershell.exe"),
    saveAfterPush: c.get<boolean>("saveAfterPush", true),
    includeDocumentModules: c.get<boolean>("includeDocumentModules", true),
    pushBeforeRun: c.get<boolean>("pushBeforeRun", true),
    autoPushOnSave: c.get<boolean>("autoPushOnSave", false),
    autoPushSaveOffice: c.get<boolean>("autoPushSaveOffice", false),
  };
}

export function bridgeOptions(cfg: PerfectVbaConfig): BridgeOptions {
  return { powerShellPath: cfg.powerShellPath };
}
