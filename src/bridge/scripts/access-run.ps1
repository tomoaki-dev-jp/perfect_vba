# access-run.ps1 — 編集済みのマクロ(Sub/Function)を Access 上で実行する。
# pull/push と異なり、実行結果を確認できるよう開いた DB は閉じずに残す。
. "$PSScriptRoot\common.ps1"

# 実行後の後始末: 可視実行は残す。ヘッドレスかつ自分で起動した場合のみ閉じる。
function Close-RunAccess($ctx, $headless) {
  if ($ctx.Started -and $headless) {
    Close-AccessContext $ctx
  }
  else {
    Release-Com $ctx.App
  }
}

$payload = Read-Payload
$ctx = $null
try {
  $ctx = New-AccessRunContext $payload.path $payload.headless
  if ($ctx.Started) { try { $ctx.App.Visible = (-not $payload.headless) } catch {} }

  $macro = [string]$payload.macro
  # Access は Application.Run "プロシージャ名"（モジュール修飾は使わない）。
  $ret = $ctx.App.Run($macro)
  $rv = if ($null -eq $ret) { $null } else { "$ret" }

  Close-RunAccess $ctx $payload.headless
  $ctx = $null
  Write-Result @{ ran = $true; macro = $macro; returnValue = $rv }
}
catch {
  if ($ctx) { try { Close-RunAccess $ctx $payload.headless } catch {}; $ctx = $null }
  Write-FromError $_
}
