// アクティビティバーの "Perfect VBA" ツリー（プロジェクト → コンポーネント）。
import * as vscode from "vscode";
import * as path from "path";
import { listProjects, ProjectEntry } from "../commands/shared";
import { ComponentRecord } from "../model/manifest";

export class ProjectNode extends vscode.TreeItem {
  readonly nodeType = "project";
  readonly projectDir: string;
  readonly manifestSource: string;

  constructor(public readonly entry: ProjectEntry) {
    super(entry.manifest.projectName, vscode.TreeItemCollapsibleState.Collapsed);
    this.projectDir = entry.dir;
    this.manifestSource = entry.manifest.source;
    this.description = path.basename(entry.manifest.source);
    this.tooltip = `${entry.manifest.source}\n取り込み: ${entry.manifest.pulledAt}`;
    this.contextValue = "vbaProject";
    this.iconPath = new vscode.ThemeIcon(
      entry.manifest.app === "access" ? "database" : "table"
    );
  }
}

export class ComponentNode extends vscode.TreeItem {
  readonly nodeType = "component";

  constructor(public readonly dir: string, public readonly rec: ComponentRecord) {
    super(rec.name, vscode.TreeItemCollapsibleState.None);
    this.description = kindLabel(rec.kind);
    const file = vscode.Uri.file(path.join(dir, rec.file));
    this.resourceUri = file;
    this.command = { command: "vscode.open", title: "開く", arguments: [file] };
    this.contextValue = "vbaComponent";
    this.iconPath = iconForKind(rec.kind);
  }
}

type Node = ProjectNode | ComponentNode;

export class PerfectVbaTreeProvider implements vscode.TreeDataProvider<Node> {
  private readonly _onDidChange = new vscode.EventEmitter<Node | undefined | void>();
  readonly onDidChangeTreeData = this._onDidChange.event;

  refresh(): void {
    this._onDidChange.fire();
  }

  getTreeItem(element: Node): vscode.TreeItem {
    return element;
  }

  async getChildren(element?: Node): Promise<Node[]> {
    if (!element) {
      const projects = await listProjects();
      return projects.map((p) => new ProjectNode(p));
    }
    if (element instanceof ProjectNode) {
      return element.entry.manifest.components
        .slice()
        .sort((a, b) => a.name.localeCompare(b.name))
        .map((rec) => new ComponentNode(element.entry.dir, rec));
    }
    return [];
  }
}

function kindLabel(kind: string): string {
  switch (kind) {
    case "std":
      return "標準モジュール";
    case "class":
      return "クラス";
    case "form":
      return "フォーム";
    case "document":
      return "ドキュメント";
    case "accForm":
      return "フォーム";
    case "accReport":
      return "レポート";
    case "accMacro":
      return "マクロ";
    default:
      return kind;
  }
}

function iconForKind(kind: string): vscode.ThemeIcon {
  switch (kind) {
    case "std":
      return new vscode.ThemeIcon("symbol-module");
    case "class":
      return new vscode.ThemeIcon("symbol-class");
    case "form":
    case "accForm":
      return new vscode.ThemeIcon("window");
    case "document":
      return new vscode.ThemeIcon("file-code");
    case "accReport":
      return new vscode.ThemeIcon("output");
    case "accMacro":
      return new vscode.ThemeIcon("run");
    default:
      return new vscode.ThemeIcon("symbol-file");
  }
}
