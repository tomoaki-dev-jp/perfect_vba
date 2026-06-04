# excel-list.ps1 — Excel ブックの VBA コンポーネント一覧を返す。
. "$PSScriptRoot\common.ps1"

$payload = Read-Payload
$ctx = $null
try {
  $ctx = New-ExcelContext $payload.path $payload.headless
  $proj = Get-ExcelVBProject $ctx
  $comps = @()
  foreach ($c in $proj.VBComponents) {
    $kind = Get-ComponentKind $c
    $comps += @{ name = $c.Name; kind = $kind; hasFrx = ($c.Type -eq $VBEXT_FORM) }
  }
  $name = $proj.Name
  Close-ExcelContext $ctx $false; $ctx = $null
  Write-Result @{ app = 'excel'; projectName = $name; components = @($comps) }
}
catch {
  if ($ctx) { Close-ExcelContext $ctx $false }
  Write-FromError $_
}
