# access-push.ps1 — 編集済みコンポーネントを Access DB へ書き戻す。
# 標準/クラス: Remove+Import。フォーム/レポート/マクロ: LoadFromText（全体上書き）。
. "$PSScriptRoot\common.ps1"

$payload = Read-Payload
$ctx = $null
$tmp = $null
try {
  $tmp = New-TempDir
  $ctx = New-AccessContext $payload.path $payload.headless ([bool]$payload.attachOnly)
  if ($ctx.NotOpen) {
    # attachOnly で対象 DB が開かれていない → 起動せずスキップ
    Close-AccessContext $ctx; $ctx = $null
    Write-Result @{ pushed = 0; skipped = 'NOT_OPEN' }
  }
  $proj = Get-AccessVBProject $ctx

  $count = 0
  foreach ($comp in $payload.components) {
    if ($comp.kind -eq 'std' -or $comp.kind -eq 'class') {
      $ext = if ($comp.fileExt) { $comp.fileExt } else { Get-KindExt $comp.kind }
      $base = Join-Path $tmp ($comp.name + $ext)
      [System.IO.File]::WriteAllBytes($base, [Convert]::FromBase64String($comp.contentB64))
      Import-VBComponent $proj $comp.name $base
    }
    elseif ($comp.kind -eq 'accForm' -or $comp.kind -eq 'accReport' -or $comp.kind -eq 'accMacro') {
      $file = Join-Path $tmp ($comp.kind + '_' + $comp.name + '.txt')
      [System.IO.File]::WriteAllBytes($file, [Convert]::FromBase64String($comp.contentB64))
      $ctx.App.LoadFromText((Get-AcType $comp.kind), $comp.name, $file)
    }
    else {
      throw "Access では未対応の種別です: $($comp.kind)"
    }
    $count++
  }

  # 開いた DB はクローズ時にコミットされる
  Close-AccessContext $ctx; $ctx = $null
  Write-Result @{ pushed = $count }
}
catch {
  if ($ctx) { Close-AccessContext $ctx }
  Write-FromError $_
}
finally {
  Remove-TempDir $tmp
}
