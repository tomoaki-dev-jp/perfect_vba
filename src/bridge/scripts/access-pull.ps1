# access-pull.ps1 — Access DB の全コンポーネントを取り出す。
# 標準/クラス: VBE Export。フォーム/レポート/マクロ: SaveAsText（定義＋コード全体）。
. "$PSScriptRoot\common.ps1"

$payload = Read-Payload
$ctx = $null
$tmp = $null
try {
  $tmp = New-TempDir
  $ctx = New-AccessContext $payload.path $payload.headless
  $proj = Get-AccessVBProject $ctx

  $comps = @()

  # 標準/クラスモジュール（VBE）
  foreach ($c in $proj.VBComponents) {
    $kind = Get-ComponentKind $c
    if ($kind -ne 'std' -and $kind -ne 'class') { continue }
    $ext = Get-KindExt $kind
    $base = Join-Path $tmp ($c.Name + $ext)
    $c.Export($base)
    $comps += @{
      name       = $c.Name
      kind       = $kind
      contentB64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($base))
    }
  }

  # フォーム/レポート/マクロ（SaveAsText）
  $objs = @()
  foreach ($f in $ctx.App.CurrentProject.AllForms) { $objs += @{ name = $f.Name; kind = 'accForm' } }
  foreach ($r in $ctx.App.CurrentProject.AllReports) { $objs += @{ name = $r.Name; kind = 'accReport' } }
  foreach ($m in $ctx.App.CurrentProject.AllMacros) { $objs += @{ name = $m.Name; kind = 'accMacro' } }
  foreach ($o in $objs) {
    $file = Join-Path $tmp ($o.kind + '_' + $o.name + '.txt')
    $ctx.App.SaveAsText((Get-AcType $o.kind), $o.name, $file)
    $comps += @{
      name       = $o.name
      kind       = $o.kind
      contentB64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file))
    }
  }

  $name = $proj.Name
  Close-AccessContext $ctx; $ctx = $null
  Write-Result @{ app = 'access'; projectName = $name; components = @($comps) }
}
catch {
  if ($ctx) { Close-AccessContext $ctx }
  Write-FromError $_
}
finally {
  Remove-TempDir $tmp
}
