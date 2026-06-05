# 変更履歴 (Changelog)

このプロジェクトの主な変更点を記録します。
フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に準拠し、
バージョニングは [Semantic Versioning](https://semver.org/lang/ja/) に従います。

## [Unreleased]

## [0.3.0] - 2026-06-05

### Added

- **保存時の自動 Push（リアルタイム反映）**: `.bas`/`.cls` を保存すると、開いている Excel / Access の VBProject へ自動で書き戻す。`files.autoSave` と併用すると実質リアルタイムに反映できる。
  - 既定はオフ。設定 `perfectVba.autoPushOnSave`、またはステータスバーの **VBA 自動Push** / コマンド **「保存時の自動 Push（リアルタイム反映）の切替」** で切り替え。
  - 対象が開かれていなければ Office を起動せずスキップする（`attachOnly`）。状態はステータスバーに集約し、保存ごとの通知は出さない。
  - 設定 `perfectVba.autoPushSaveOffice`（既定オフ）で、自動 Push のたびに Office ファイルを保存するか選べる。既定では起動中の VBProject にコードのみ即反映する。
  - 安全のため、自動 Push の対象は標準モジュール・クラス・ドキュメントモジュールに限定（フォームや Access のフォーム/レポート/マクロのような全体上書きは従来どおり手動 Push）。

### Changed

- `Push` の進捗表示を呼び出し元で選べるように内部リファクタ（通知 / ウィンドウ / 非表示）。既存コマンドの表示は従来どおり。

## [0.2.0] - 2026-06-03

最初の公開リリース。

### Added

- **Pull**: Office ファイル（Excel / Access）から全 VBA コンポーネントをワークスペースへ取り込み
- **Push**: 編集したファイルを Office ファイルへ書き戻し（保存まで）
- **Run**: `.bas` の各 `Sub`/`Function` 上の **▶ 実行** CodeLens、コマンド、ツリーからマクロを実行（書き戻し → `Application.Run`）
- **このファイルを Push**: アクティブな `.bas`/`.cls`/`.frm` を単体で書き戻し
- **再取り込み (Refresh)** / **一覧の更新** コマンド
- **紐づく Office ファイルを開く** コマンド
- **VBA プロジェクトへのアクセスを信頼する設定の有効化** コマンド
- Excel: 標準モジュール(.bas) / クラス(.cls) / ユーザーフォーム(.frm + .frx) / ドキュメントモジュール(ThisWorkbook・Sheet) に対応
- Access: 標準・クラスモジュール / フォーム・レポート（コードビハインド込み）/ マクロに対応
- 実行中の Excel / Access インスタンスへの自動接続（未起動時はバックグラウンド起動）
- 日本語(Shift_JIS)対応：取り込み時に UTF-8 へ変換し、書き戻し時に元のエンコーディングへ復元
- ツリービュー / ステータスバー / コマンドパレットからの操作
- 設定項目: `workspaceRoot` / `encoding` / `headless` / `powerShellPath` / `saveAfterPush` / `includeDocumentModules` / `pushBeforeRun`

