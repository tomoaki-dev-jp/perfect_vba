# Perfect VBA

**Excel / Access の VBA を VSCode 上で取り込み（Pull）・編集し、Office ファイルへ書き戻す（Push）** VSCode 拡張機能です。

VBA は標準では古い VBE エディタでしか編集できませんが、本拡張を使うと VSCode の編集体験（補完・検索・Git 管理）をそのまま VBA に活かせます。**Excel と Access の両方を双方向でサポート**します。

---

## 主な機能

- **Pull**: Office ファイルから全 VBA コンポーネントを取り出し、ワークスペースにテキストファイルとして展開
- **Push**: 編集したファイルを Office ファイルへ書き戻し（保存まで）
- **Run**: 編集したマクロを VSCode から実行。`.bas` の各 `Sub`/`Function` の上に出る **▶ 実行** ボタン（または ▶ ボタン / コマンド）で、書き戻し → `Application.Run` 実行までを一気に行う
- **対応コンポーネント**
  - Excel: 標準モジュール(.bas) / クラス(.cls) / ユーザーフォーム(.frm + .frx) / ドキュメントモジュール(ThisWorkbook・Sheet)
  - Access: 標準・クラスモジュール / フォーム・レポート（コードビハインド込み）/ マクロ
- **実行中インスタンスに自動接続**：対象を Excel/Access で開いていればそのインスタンスへ、閉じていればバックグラウンドで起動して処理
- **保存時の自動 Push（リアルタイム反映）**：`.bas`/`.cls` を保存すると、開いている Office の VBProject へ即書き戻し。`files.autoSave` と併用すれば実質リアルタイムに反映（既定オフ・オプトイン）
- **日本語(Shift_JIS)対応**：VBE が出力する Shift_JIS を UTF-8 に変換して編集、書き戻し時に元のエンコーディングへ復元
- ツリービュー / ステータスバー / コマンドパレットから操作

---

## 動作環境・前提

- **Windows**（COM 自動化を使用するため Windows 専用）
- **Microsoft Excel / Access** がインストール済み
- **PowerShell**（Windows 標準の `powershell.exe` 5.1 で動作）
- **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定が有効**であること
  - 各アプリの [ファイル] → [オプション] → [トラスト センター] → [トラスト センターの設定] → [マクロの設定] で有効化
  - または本拡張のコマンド **「Perfect VBA: VBA プロジェクトへのアクセスを信頼する設定を有効化」** を実行（設定後は対象アプリの再起動が必要）

---

## 使い方

1. 任意のフォルダ（ワークスペース）を VSCode で開く
2. コマンドパレット（`Ctrl+Shift+P`）で **「Perfect VBA: Office ファイルから取り込み (Pull)」** を実行し、`.xlsm` / `.accdb` などを選択
   - エクスプローラで Office ファイルを右クリック →「取り込み」でも可
3. `.vba/<ファイル名>/` 配下に `.bas` / `.cls` / `.frm` などが展開される
4. VSCode で編集して保存
5. **「Perfect VBA: Office ファイルへ書き戻し (Push)」**、ツリーの Push ボタン、またはエディタ右上 / ステータスバーの **VBA Push** で書き戻し

### マクロを実行する（Run）

編集したマクロをそのまま VSCode から実行できます。

- `.bas`（標準モジュール）を開くと、引数なしの `Public` な `Sub`/`Function` の上に **▶ 実行** ボタン（CodeLens）が表示されます。クリックすると **そのファイルを Push してから当該マクロを実行** します。
- エディタ右上の **▶** ボタン、コマンドパレットの **「Perfect VBA: マクロを実行 (Run)」**、ツリーのプロジェクト項目の **▶** からも実行できます（実行するマクロを一覧から選択）。
- `Function` の場合は戻り値を通知に表示します。実行結果を確認できるよう、拡張機能が開いた Excel / Access は **閉じずに表示したまま** 残します。
- 既定では実行前に自動 Push します（`perfectVba.pushBeforeRun`）。オフにすると Office 側の現在のコードを実行します。

### リアルタイム反映（保存時の自動 Push）

Excel / Access で対象ファイル（VBE / マクロ）を開いたまま、VSCode で編集 → **保存するたびに、開いているインスタンスへ自動で Push** できます。VBE のコードウィンドウを開いたままでも `CodeModule` 直接置換で即時反映されるため、**書き戻しのたびに Office を開き直す必要はありません**。

- 有効化：コマンド **「Perfect VBA: 保存時の自動 Push（リアルタイム反映）の切替」**、ステータスバーの **`$(sync) VBA 自動Push`** をクリック、または設定 `perfectVba.autoPushOnSave` を `true`。
- **実質リアルタイムにするコツ**：VSCode の `files.autoSave` を `afterDelay`（例: `"files.autoSaveDelay": 300`）にすると、入力が止まるたびに保存 → 自動 Push され、ほぼ打ち込みながら反映できます。
- 対象が**開かれていなければ Office は起動せずスキップ**します（保存のたびに勝手に Excel/Access が立ち上がることはありません）。状態はステータスバー（`$(sync~spin)` Push 中 / `$(check)` 同期 / `$(circle-slash)` 未起動 / `$(error)` 失敗）に表示し、保存ごとの通知ポップアップは出しません。
- 既定では**起動中の VBProject にコードのみ即反映**し、Office ファイル自体は保存しません（連続保存での負荷・競合を避けるため）。毎回ファイルも保存したい場合は `perfectVba.autoPushSaveOffice` を `true` にしてください。
- 安全のため、自動 Push の対象は**標準モジュール(.bas) / クラス(.cls) / ドキュメントモジュール**に限定しています。ユーザーフォームや Access のフォーム/レポート/マクロ（オブジェクト全体の上書きになるもの）は、従来どおり**手動 Push** で書き戻してください。

### コマンド一覧

| コマンド | 説明 |
| --- | --- |
| `Perfect VBA: Office ファイルから取り込み (Pull)` | Office ファイル → ワークスペース |
| `Perfect VBA: Office ファイルへ書き戻し (Push)` | ワークスペース → Office ファイル |
| `Perfect VBA: このファイルを Push` | アクティブな .bas/.cls 等を単体で書き戻し |
| `Perfect VBA: マクロを実行 (Run)` | （必要なら Push してから）マクロを `Application.Run` で実行 |
| `Perfect VBA: 再取り込み (Refresh)` | 取り込み直し |
| `Perfect VBA: VBA プロジェクトへのアクセスを信頼する設定を有効化` | Trust 設定をレジストリへ書き込み |
| `Perfect VBA: 紐づく Office ファイルを開く` | 取り込み元ファイルを既定アプリで開く |
| `Perfect VBA: 保存時の自動 Push（リアルタイム反映）の切替` | 保存時の自動 Push を On/Off |

### 設定

| 設定キー | 既定 | 説明 |
| --- | --- | --- |
| `perfectVba.workspaceRoot` | `.vba` | 取り込み先のルートフォルダ |
| `perfectVba.encoding` | `auto` | `auto`(システム ANSI=日本語環境では Shift_JIS) / `shift_jis` / `utf-8` |
| `perfectVba.headless` | `false` | 拡張機能が起動する Office を非表示にする |
| `perfectVba.powerShellPath` | `powershell.exe` | 使用する PowerShell |
| `perfectVba.saveAfterPush` | `true` | Push 後にファイルを保存 |
| `perfectVba.includeDocumentModules` | `true` | ThisWorkbook/Sheet 等も対象にする |
| `perfectVba.pushBeforeRun` | `true` | Run の前に対象を自動 Push（保存）する |
| `perfectVba.autoPushOnSave` | `false` | `.bas`/`.cls` の保存時に、開いている Office へ自動 Push（閉じていればスキップ）。`files.autoSave` 併用で実質リアルタイム |
| `perfectVba.autoPushSaveOffice` | `false` | 自動 Push のたびに Office ファイルも保存する（既定はコードのみ即反映） |

---

## 仕組み（概要）

```
VSCode 拡張(TypeScript)
  └─ 子プロセス起動 → PowerShell(.ps1) ── COM ──→ Excel / Access (VBIDE API)
        stdin: payload(JSON)            stdout: 結果(JSON, 生バイトは base64)
```

- 標準/クラス/フォームは VBE の `Export` / `Import`（同名は `Remove` してから `Import`）
- ThisWorkbook/Sheet 等のドキュメントモジュールは Remove/Import 不可のため `CodeModule` の行置換
- Access のフォーム/レポート/マクロは `SaveAsText` / `LoadFromText`
- エンコーディング変換は拡張機能側（`iconv-lite`）に集約。`.frx` はバイナリのまま保持

---

## 既知の制限

- `.frx`（ユーザーフォームのレイアウト・画像）はバイナリのため、**フォームのコードは編集できますがレイアウト変更は VBE 側**で行ってください。
- Access の `LoadFromText` はオブジェクト全体を**確認なしで上書き**します。重要な DB は Push 前にバックアップしてください。
- **Run の対象**は「標準モジュール(.bas)内の、引数なしで `Private` でない `Sub`/`Function`」です（Excel の Alt+F8 で実行できるものと同じ範囲）。クラス/ドキュメントモジュールのメソッドや、必須引数のあるマクロは一覧に出ません。
- Run は対象ブック/DB を**マクロ有効**で開くため、`Workbook_Open` などの自動実行マクロが動作することがあります。`Debug.Print` の出力（イミディエイト ウィンドウ）は VSCode 側では取得できません。`Function` の戻り値のみ通知に表示します。
- Microsoft は Office の完全な無人自動化を公式サポートしていません。既定では可視で起動し、拡張機能が開いたインスタンスのみ処理後に閉じます（ユーザーが開いていたものは閉じません）。
- パスワード保護された VBA プロジェクトは対象外です。
- Access の SaveAsText 出力エンコーディングは Office のバージョンにより異なる場合があります（BOM を自動判定しますが、フォーム/レポートの往復はベストエフォートです）。

---

## 開発

```bash
npm install
npm run compile      # esbuild バンドル + tsc 型チェック
npm run test:unit    # ユニットテスト（Office 不要）
```

VSCode でこのフォルダを開き **F5**（Extension Development Host）でデバッグ起動できます。

### 手動の往復テスト

```powershell
# テスト用 .xlsm を生成（要 Excel + Trust 設定）
powershell -ExecutionPolicy Bypass -File sample/make-test-book.ps1
```

生成された `sample/Test.xlsm` を Pull → 編集 → Push し、VBE で反映を確認します。

---

## 変更履歴

リリースごとの変更点は [CHANGELOG.md](CHANGELOG.md) を参照してください。

## ライセンス

[MIT License](LICENSE) © 2026 tomo
