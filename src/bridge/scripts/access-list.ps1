# access-list.ps1 — Access DB のコンポーネント一覧。
# 標準/クラスモジュールは VBE、フォーム/レポート/マクロは CurrentProject から列挙。
. "$PSScriptRoot\common.ps1"

$payload = Read-Payload
$ctx = $null
try {
  $ctx = New-AccessContext $payload.path $payload.headless
  $proj = Get-AccessVBProject $ctx

  $comps = @()
  foreach ($c in $proj.VBComponents) {
    $kind = Get-ComponentKind $c
    if ($kind -eq 'std' -or $kind -eq 'class') {
      $comps += @{ name = $c.Name; kind = $kind }
    }
    # document(Form_/Report_ のコードビハインド)は SaveAsText 側で全体を扱うため除外
  }
  foreach ($f in $ctx.App.CurrentProject.AllForms) { $comps += @{ name = $f.Name; kind = 'accForm' } }
  foreach ($r in $ctx.App.CurrentProject.AllReports) { $comps += @{ name = $r.Name; kind = 'accReport' } }
  foreach ($m in $ctx.App.CurrentProject.AllMacros) { $comps += @{ name = $m.Name; kind = 'accMacro' } }

  $name = $proj.Name
  Close-AccessContext $ctx; $ctx = $null
  Write-Result @{ app = 'access'; projectName = $name; components = @($comps) }
}
catch {
  if ($ctx) { Close-AccessContext $ctx }
  Write-FromError $_
}
