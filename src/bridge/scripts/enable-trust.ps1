# enable-trust.ps1 — 「VBA プロジェクトオブジェクトモデルへのアクセスを信頼する」を有効化。
# HKCU の各 Office バージョン配下に <App>\Security\AccessVBOM = 1 を書き込む。
# 反映には対象アプリの再起動が必要。
. "$PSScriptRoot\common.ps1"

$payload = Read-Payload
try {
  $appName = if ($payload.app -eq 'access') { 'Access' } else { 'Excel' }
  $base = 'HKCU:\Software\Microsoft\Office'
  $versions = @()

  $verKeys = @()
  if (Test-Path $base) {
    $verKeys = Get-ChildItem $base | Where-Object { $_.PSChildName -match '^\d+\.\d+$' }
  }
  if (-not $verKeys -or $verKeys.Count -eq 0) {
    # 既知の代表バージョンにフォールバック
    $verKeys = @('16.0', '15.0', '14.0') | ForEach-Object { [pscustomobject]@{ PSChildName = $_; PSPath = (Join-Path $base $_) } }
  }

  foreach ($v in $verKeys) {
    $secPath = Join-Path (Join-Path ("HKCU:\Software\Microsoft\Office\" + $v.PSChildName) $appName) 'Security'
    try {
      if (-not (Test-Path $secPath)) { New-Item -Path $secPath -Force | Out-Null }
      New-ItemProperty -Path $secPath -Name 'AccessVBOM' -Value 1 -PropertyType DWord -Force | Out-Null
      $versions += $v.PSChildName
    }
    catch {}
  }

  Write-Result @{ app = $payload.app; appName = $appName; versions = $versions }
}
catch {
  Write-FromError $_
}
