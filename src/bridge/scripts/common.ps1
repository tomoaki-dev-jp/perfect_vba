# common.ps1 — Perfect VBA の PowerShell COM ブリッジ共通基盤。
# 各操作スクリプトから dot-source して使う。
# 入出力は UTF-8 / 単一 JSON。エラーも JSON(ok=$false)で返し、exit code は 0 を基本とする。

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# stdin/stdout を UTF-8 に固定（日本語コードや日本語パスを正しく授受するため）
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}
try { [Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}

# --- COM 型定数 ---
$script:VBEXT_STD = 1
$script:VBEXT_CLASS = 2
$script:VBEXT_FORM = 3
$script:VBEXT_DOCUMENT = 100
# Access AcObjectType
$script:AC_FORM = 2
$script:AC_REPORT = 3
$script:AC_MACRO = 4

# stdin を生バイトで読み、UTF-8 として解釈する（リダイレクト時の文字化け対策）。
function Read-Payload {
  $stdin = [Console]::OpenStandardInput()
  $ms = New-Object System.IO.MemoryStream
  $stdin.CopyTo($ms)
  $bytes = $ms.ToArray()
  $ms.Dispose()
  if ($bytes.Length -eq 0) { return [pscustomobject]@{} }
  $raw = [System.Text.Encoding]::UTF8.GetString($bytes)
  if ([string]::IsNullOrWhiteSpace($raw)) { return [pscustomobject]@{} }
  return ($raw | ConvertFrom-Json)
}

function Get-Acp {
  try { return [int][System.Text.Encoding]::Default.CodePage } catch { return 0 }
}

# JSON を UTF-8 バイトとして標準出力へ直接書き出す（コンソールエンコーディングに依存しない）。
function Write-OutJson($obj) {
  $json = $obj | ConvertTo-Json -Depth 30 -Compress
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  $out = [Console]::OpenStandardOutput()
  $out.Write($bytes, 0, $bytes.Length)
  $out.Flush()
}

function Write-Result($data) {
  Write-OutJson @{ ok = $true; data = $data; acp = (Get-Acp) }
  exit 0
}

function Write-ErrResult($code, $message) {
  Write-OutJson @{ ok = $false; error = @{ code = $code; message = "$message" }; acp = (Get-Acp) }
  exit 0
}

function Write-FromError($err) {
  $msg = if ($err.Exception) { $err.Exception.Message } else { "$err" }
  if ($msg -match 'trust' -or $msg -match '信頼' -or $msg -match 'プロジェクトへのアクセス' -or $msg -match '0x800A03EC') {
    Write-ErrResult 'TRUST_DISABLED' $msg
  }
  elseif ($msg -match '排他' -or $msg -match 'exclusive') {
    Write-ErrResult 'EXCLUSIVE_LOCKED' $msg
  }
  else {
    Write-ErrResult 'COM_ERROR' $msg
  }
}

function Resolve-OfficePath($p) {
  return [System.IO.Path]::GetFullPath($p)
}

function New-TempDir {
  $dir = Join-Path ([System.IO.Path]::GetTempPath()) ("perfectvba_" + [System.Guid]::NewGuid().ToString('N'))
  [void][System.IO.Directory]::CreateDirectory($dir)
  return $dir
}

function Remove-TempDir($dir) {
  if ($dir -and (Test-Path $dir)) {
    try { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
  }
}

function Release-Com($o) {
  if ($null -ne $o) {
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {}
  }
}

# ---------------- Excel ----------------

function New-ExcelContext($path, $headless) {
  $full = Resolve-OfficePath $path
  $app = $null; $started = $false; $openedDoc = $false; $origEvents = $null
  try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $app = $null }
  if (-not $app) {
    $app = New-Object -ComObject Excel.Application
    $started = $true
    try { $app.Visible = (-not $headless) } catch {}
  }

  # 既に開いているブックを探す（ユーザーのインスタンスを尊重）
  $wb = $null
  foreach ($w in $app.Workbooks) {
    try { if ($w.FullName -ieq $full) { $wb = $w; break } } catch {}
  }

  if (-not $wb) {
    # 自分で開く場合のみ、安全フラグを立てる（attach 時は後で元に戻す）
    if (-not $started) { try { $origEvents = $app.EnableEvents } catch {} }
    try { $app.DisplayAlerts = $false } catch {}
    try { $app.EnableEvents = $false } catch {}
    try { $app.AutomationSecurity = 3 } catch {}  # msoAutomationSecurityForceDisable: マクロ自動実行を抑止
    if (-not (Test-Path $full)) { throw "ファイルが見つかりません: $full" }
    $wb = $app.Workbooks.Open($full)
    $openedDoc = $true
  }
  return [pscustomobject]@{ App = $app; Doc = $wb; Started = $started; OpenedDoc = $openedDoc; Full = $full; OrigEvents = $origEvents }
}

function Close-ExcelContext($ctx, $save) {
  if (-not $ctx) { return }
  try { if ($save -and $ctx.Doc) { $ctx.Doc.Save() } } catch {}
  try { if ($ctx.OpenedDoc -and $ctx.Doc) { $ctx.Doc.Close($false) } } catch {}
  # attach したインスタンスでイベントを止めていた場合は元に戻す
  if (-not $ctx.Started -and $null -ne $ctx.OrigEvents) {
    try { $ctx.App.EnableEvents = $ctx.OrigEvents } catch {}
  }
  try { if ($ctx.Started -and $ctx.App) { $ctx.App.Quit() } } catch {}
  Release-Com $ctx.Doc
  Release-Com $ctx.App
}

function Get-ExcelVBProject($ctx) {
  # Trust 未許可ならここで例外 -> 呼び出し側 catch で TRUST_DISABLED に写像
  $proj = $ctx.Doc.VBProject
  if ($null -eq $proj) { throw "VBProject にアクセスできません（信頼設定を確認してください）。" }
  return $proj
}

# ---------------- Access ----------------

function New-AccessContext($path, $headless) {
  $full = Resolve-OfficePath $path
  $app = $null; $started = $false; $openedDoc = $false
  try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Access.Application') } catch { $app = $null }

  if ($app) {
    $cur = $null
    try { $cur = $app.CurrentProject.FullName } catch { $cur = $null }
    if ($cur -and ($cur -ieq $full)) {
      # ユーザーが開いている当該 DB にそのまま接続
      return [pscustomobject]@{ App = $app; Started = $false; OpenedDoc = $false; Full = $full }
    }
    # 別の DB を開いている既存インスタンスは使わず、新規インスタンスを起動する
    Release-Com $app
    $app = $null
  }

  $app = New-Object -ComObject Access.Application
  $started = $true
  try { $app.Visible = (-not $headless) } catch {}
  if (-not (Test-Path $full)) { throw "ファイルが見つかりません: $full" }
  try { $app.AutomationSecurity = 3 } catch {}
  # 排他オープン（SaveAsText/LoadFromText に必要）
  $app.OpenCurrentDatabase($full, $true)
  $openedDoc = $true
  return [pscustomobject]@{ App = $app; Started = $started; OpenedDoc = $openedDoc; Full = $full }
}

function Close-AccessContext($ctx) {
  if (-not $ctx) { return }
  try { if ($ctx.OpenedDoc -and $ctx.App) { $ctx.App.CloseCurrentDatabase() } } catch {}
  try { if ($ctx.Started -and $ctx.App) { $ctx.App.Quit() } } catch {}
  Release-Com $ctx.App
}

function Get-AccessVBProject($ctx) {
  $full = $ctx.Full
  $vbe = $ctx.App.VBE
  $found = $null
  foreach ($p in $vbe.VBProjects) {
    try { if ($p.FileName -ieq $full) { $found = $p; break } } catch {}
  }
  if (-not $found) {
    try { $found = $vbe.ActiveVBProject } catch {}
  }
  if (-not $found) { $found = $vbe.VBProjects.Item(1) }
  if ($null -eq $found) { throw "Access の VBProject にアクセスできません（信頼設定を確認してください）。" }
  return $found
}

# ---------------- 共通: コンポーネント種別 ----------------

function Get-ComponentKind($comp) {
  switch ($comp.Type) {
    1 { 'std' }
    2 { 'class' }
    3 { 'form' }
    100 { 'document' }
    default { 'std' }
  }
}

function Get-KindExt($kind) {
  switch ($kind) {
    'std' { '.bas' }
    'class' { '.cls' }
    'form' { '.frm' }
    default { '.bas' }
  }
}

function Find-Component($proj, $name) {
  foreach ($c in $proj.VBComponents) {
    if ($c.Name -eq $name) { return $c }
  }
  return $null
}

# エクスポートされた .bas/.cls テキストから、モジュールヘッダを除いたコード本体を返す。
# クラスの "VERSION 1.0 CLASS" + "BEGIN ... END" ブロックと、先頭の "Attribute VB_*" 行を
# 取り除く。CodeModule.AddFromString はこれらのヘッダ行を受け付けない（コードとして混入する）ため。
function Get-CodeBody($text) {
  if ($null -eq $text) { return '' }
  $lines = [regex]::Split([string]$text, "`r`n|`n|`r")
  $i = 0
  if ($i -lt $lines.Count -and $lines[$i] -match '^\s*VERSION\s') {
    $i++
    if ($i -lt $lines.Count -and $lines[$i] -match '^\s*BEGIN\b') {
      while ($i -lt $lines.Count -and $lines[$i] -notmatch '^\s*END\s*$') { $i++ }
      if ($i -lt $lines.Count) { $i++ }  # END 行を読み飛ばす
    }
  }
  while ($i -lt $lines.Count -and $lines[$i] -match '^\s*Attribute\s+VB_') { $i++ }
  if ($i -ge $lines.Count) { return '' }
  return ($lines[$i..($lines.Count - 1)] -join "`r`n")
}

# 既存の std/class モジュールのコードを CodeModule 直接操作で丸ごと置き換える。
# Remove+Import と違い VBE でコードウィンドウを開いたままでも確実に即時反映され、
# 遅延削除による重複モジュール(Module11 等)も生じない。
function Set-ModuleCodeInPlace($comp, $exportText) {
  $cm = $comp.CodeModule
  if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
  $body = Get-CodeBody $exportText
  if ($body -and $body.Length -gt 0) { $cm.AddFromString($body) }
}

# 標準/クラス/フォームを再 Import する。同名既存は Remove してから取り込む（重複回避）。
function Import-VBComponent($proj, $name, $file) {
  $existing = Find-Component $proj $name
  if ($existing) {
    if ($existing.Type -eq 100) { throw "ドキュメントモジュールは Import できません: $name" }
    $proj.VBComponents.Remove($existing)
    Start-Sleep -Milliseconds 50
  }
  [void]$proj.VBComponents.Import($file)
}

# Access SaveAsText/LoadFromText の AcObjectType 番号を kind から得る。
function Get-AcType($kind) {
  switch ($kind) {
    'accForm' { $script:AC_FORM }
    'accReport' { $script:AC_REPORT }
    'accMacro' { $script:AC_MACRO }
    default { throw "Access オブジェクト種別が不正です: $kind" }
  }
}

# ドキュメントモジュール(ThisWorkbook/Sheet 等)のコードを行置換で反映する。
function Set-DocumentCode($proj, $name, $code) {
  $c = Find-Component $proj $name
  if (-not $c) { throw "コンポーネントが見つかりません: $name" }
  $cm = $c.CodeModule
  if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
  if ($code -and $code.Length -gt 0) { $cm.AddFromString($code) }
}

# ---------------- 実行(Run)用コンテキスト ----------------
# pull/push と異なり、(1)マクロを有効化して開き、(2)結果をユーザーが確認できるよう
# 原則ウィンドウを残す（呼び出し側で Close しない）。自分で起動したインスタンスにのみ
# 可視/セキュリティを設定し、ユーザーが既に開いているインスタンスには手を加えない。

function New-ExcelRunContext($path, $headless) {
  $full = Resolve-OfficePath $path
  $app = $null; $started = $false; $openedDoc = $false
  try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $app = $null }
  if (-not $app) {
    $app = New-Object -ComObject Excel.Application
    $started = $true
    try { $app.Visible = (-not $headless) } catch {}
  }

  $wb = $null
  foreach ($w in $app.Workbooks) {
    try { if ($w.FullName -ieq $full) { $wb = $w; break } } catch {}
  }
  if (-not $wb) {
    if (-not (Test-Path $full)) { throw "ファイルが見つかりません: $full" }
    try { $app.AutomationSecurity = 1 } catch {}  # msoAutomationSecurityLow: マクロ有効
    $wb = $app.Workbooks.Open($full)
    $openedDoc = $true
  }
  return [pscustomobject]@{ App = $app; Doc = $wb; Started = $started; OpenedDoc = $openedDoc; Full = $full }
}

function New-AccessRunContext($path, $headless) {
  $full = Resolve-OfficePath $path
  $app = $null; $started = $false; $openedDoc = $false
  try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Access.Application') } catch { $app = $null }

  if ($app) {
    $cur = $null
    try { $cur = $app.CurrentProject.FullName } catch { $cur = $null }
    if ($cur -and ($cur -ieq $full)) {
      return [pscustomobject]@{ App = $app; Started = $false; OpenedDoc = $false; Full = $full }
    }
    Release-Com $app
    $app = $null
  }

  $app = New-Object -ComObject Access.Application
  $started = $true
  try { $app.Visible = (-not $headless) } catch {}
  if (-not (Test-Path $full)) { throw "ファイルが見つかりません: $full" }
  try { $app.AutomationSecurity = 1 } catch {}  # マクロ有効
  # 実行用は非排他で開く（読み取り中心のため、可能な範囲でユーザーと共存）
  $app.OpenCurrentDatabase($full, $false)
  $openedDoc = $true
  return [pscustomobject]@{ App = $app; Started = $started; OpenedDoc = $openedDoc; Full = $full }
}
