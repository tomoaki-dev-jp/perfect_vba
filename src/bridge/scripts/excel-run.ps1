# excel-run.ps1 — 編集済みのマクロ(Sub/Function)を Excel 上で実行する。
# pull/push と異なり、実行結果を確認できるよう開いたブックは閉じずに残す。
. "$PSScriptRoot\common.ps1"

# 実行後の後始末: 可視実行は結果を残す。ヘッドレスかつ自分で起動した場合のみ閉じる。
function Close-RunExcel($ctx, $headless) {
  if ($ctx.Started -and $headless) {
    Close-ExcelContext $ctx $false
  }
  else {
    Release-Com $ctx.Doc
    Release-Com $ctx.App
  }
}

$payload = Read-Payload
$ctx = $null
try {
  $ctx = New-ExcelRunContext $payload.path $payload.headless
  $doc = $ctx.Doc
  if ($ctx.Started) { try { $ctx.App.Visible = (-not $payload.headless) } catch {} }
  try { $doc.Activate() } catch {}

  $macro = [string]$payload.macro
  $ref = if ($payload.module) { "$($payload.module).$macro" } else { $macro }

  # 'Book.xlsm'!Module.Proc で実行。モジュール名は manifest 由来で確実なため単一呼び出し。
  # （bare 名へのリトライはマクロ内部エラー時に二重実行になり得るので行わない）
  $ret = $ctx.App.Run("'" + $doc.Name + "'!" + $ref)
  $rv = if ($null -eq $ret) { $null } else { "$ret" }

  Close-RunExcel $ctx $payload.headless
  $ctx = $null
  Write-Result @{ ran = $true; macro = $ref; returnValue = $rv }
}
catch {
  if ($ctx) { try { Close-RunExcel $ctx $payload.headless } catch {}; $ctx = $null }
  Write-FromError $_
}
