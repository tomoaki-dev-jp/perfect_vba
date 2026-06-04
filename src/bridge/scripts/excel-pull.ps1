# excel-pull.ps1 — Excel ブックの全 VBA コンポーネントを取り出す。
# byte 種別(std/class/form)は Export した生バイトを base64 で、document は Unicode 文字列で返す。
. "$PSScriptRoot\common.ps1"

$payload = Read-Payload
$ctx = $null
$tmp = $null
try {
  $tmp = New-TempDir
  $ctx = New-ExcelContext $payload.path $payload.headless
  $proj = Get-ExcelVBProject $ctx

  $comps = @()
  foreach ($c in $proj.VBComponents) {
    $kind = Get-ComponentKind $c
    if ($kind -eq 'document' -and -not $payload.includeDocumentModules) { continue }

    if ($kind -eq 'document') {
      $cm = $c.CodeModule
      $code = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { '' }
      $comps += @{ name = $c.Name; kind = $kind; codeText = $code }
    }
    else {
      $ext = Get-KindExt $kind
      $base = Join-Path $tmp ($c.Name + $ext)
      $c.Export($base)
      $entry = @{
        name       = $c.Name
        kind       = $kind
        contentB64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($base))
      }
      if ($kind -eq 'form') {
        $frx = [System.IO.Path]::ChangeExtension($base, '.frx')
        if (Test-Path $frx) {
          $entry.frxB64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($frx))
        }
      }
      $comps += $entry
    }
  }

  $name = $proj.Name
  Close-ExcelContext $ctx $false; $ctx = $null
  Write-Result @{ app = 'excel'; projectName = $name; components = @($comps) }
}
catch {
  if ($ctx) { Close-ExcelContext $ctx $false }
  Write-FromError $_
}
finally {
  Remove-TempDir $tmp
}
