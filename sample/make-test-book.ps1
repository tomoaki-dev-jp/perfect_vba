# make-test-book.ps1 — Perfect VBA の手動テスト用 .xlsm を生成する。
# 日本語コメント入り標準モジュール / クラス / ThisWorkbook / Sheet コード / UserForm を含む。
#
# 前提:
#   - Excel がインストール済み
#   - [トラスト センター]→[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] が有効
#     （拡張機能の「信頼設定を有効化」コマンドでも設定可能。設定後は Excel を再起動）
#
# 実行（このセッションで自分で実行する場合）:
#   ! powershell -ExecutionPolicy Bypass -File sample/make-test-book.ps1

$ErrorActionPreference = 'Stop'
$out = Join-Path $PSScriptRoot 'Test.xlsm'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Add()
  $proj = $wb.VBProject  # ここで失敗する場合は信頼設定が未有効

  # 標準モジュール（日本語コメント・日本語プロシージャ名）
  $std = $proj.VBComponents.Add(1)  # vbext_ct_StdModule
  $std.Name = 'Module1'
  $std.CodeModule.AddFromString(@"
Attribute VB_Name = "Module1"
Option Explicit

' 日本語のコメント：合計を求める
Public Function 合計(ByVal a As Long, ByVal b As Long) As Long
    合計 = a + b
End Function

Sub テスト実行()
    MsgBox "結果は " & 合計(2, 3) & " です①②③"
End Sub
"@)

  # クラスモジュール
  $cls = $proj.VBComponents.Add(2)  # vbext_ct_ClassModule
  $cls.Name = 'clsGreeter'
  $cls.CodeModule.AddFromString(@"
Option Explicit

Public Sub あいさつ()
    Debug.Print "こんにちは、世界"
End Sub
"@)

  # ThisWorkbook（ドキュメントモジュール）
  $wbComp = $proj.VBComponents('ThisWorkbook')
  $wbComp.CodeModule.AddFromString(@"
Private Sub Workbook_Open()
    ' 起動時の処理（テスト用）
End Sub
"@)

  # Sheet1（ドキュメントモジュール）
  $sheetCodeName = $wb.Sheets(1).CodeName
  $sh = $proj.VBComponents($sheetCodeName)
  $sh.CodeModule.AddFromString(@"
Private Sub Worksheet_Change(ByVal Target As Range)
    ' セル変更時（テスト用）
End Sub
"@)

  # UserForm（.frm + .frx）
  $frm = $proj.VBComponents.Add(3)  # vbext_ct_MSForm
  $frm.Name = 'UserForm1'
  $frm.CodeModule.AddFromString(@"
Private Sub UserForm_Click()
    MsgBox "フォームがクリックされました"
End Sub
"@)

  $wb.SaveAs($out, 52)  # xlOpenXMLWorkbookMacroEnabled (.xlsm)
  $wb.Close($false)
  Write-Host "作成しました: $out"
}
finally {
  $xl.Quit()
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
