# 変更履歴 (Changelog)

このプロジェクトの主な変更点を記録します。
フォーマットは [Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) に準拠し、
バージョニングは [Semantic Versioning](https://semver.org/lang/ja/) に従います。

## [Unreleased]

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

