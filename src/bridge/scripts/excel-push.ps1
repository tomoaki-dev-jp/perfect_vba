# excel-push.ps1 — 編集済みコンポーネントを Excel ブックへ書き戻す。
# std/class/form: 一時ファイルに書いて Remove+Import。document: 行置換。
. "$PSScriptRoot\common.ps1"

$payload = Read-Payload
$ctx = $null
$tmp = $null
try {
  $tmp = New-TempDir
  $ctx = New-ExcelContext $payload.path $payload.headless
  $proj = Get-ExcelVBProject $ctx

  $count = 0
  foreach ($comp in $payload.components) {
    if ($comp.kind -eq 'document') {
      Set-DocumentCode $proj $comp.name $comp.codeText
    }
    elseif ($comp.kind -eq 'form') {
      # フォームは .frm/.frx の Import が必須。同名既存は Remove してから取り込む。
      $ext = if ($comp.fileExt) { $comp.fileExt } else { Get-KindExt $comp.kind }
      $base = Join-Path $tmp ($comp.name + $ext)
      [System.IO.File]::WriteAllBytes($base, [Convert]::FromBase64String($comp.contentB64))
      if ($comp.frxB64) {
        $frx = [System.IO.Path]::ChangeExtension($base, '.frx')
        [System.IO.File]::WriteAllBytes($frx, [Convert]::FromBase64String($comp.frxB64))
      }
      Import-VBComponent $proj $comp.name $base
    }
    else {
      # std/class: 既存・新規とも CodeModule 直接操作で反映する（VBE 表示中でも確実に
      # 即時反映。Import の名前付け揺れ・クラスヘッダ依存・Remove 遅延を避ける）。
      # 新規は VBComponents.Add で作成してから名前を付ける。
      $target = Find-Component $proj $comp.name
      if (-not $target) {
        $type = if ($comp.kind -eq 'class') { $script:VBEXT_CLASS } else { $script:VBEXT_STD }
        $target = $proj.VBComponents.Add($type)
        try { $target.Name = $comp.name } catch {}
      }
      if ($target.Type -ne 100 -and $null -ne $comp.codeText) {
        Set-ModuleCodeInPlace $target $comp.codeText
      }
    }
    $count++
  }

  Close-ExcelContext $ctx $payload.save; $ctx = $null
  Write-Result @{ pushed = $count }
}
catch {
  if ($ctx) { Close-ExcelContext $ctx $false }
  Write-FromError $_
}
finally {
  Remove-TempDir $tmp
}
