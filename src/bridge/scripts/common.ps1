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
$script:AC_MODULE = 5

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

# ROT(Running Object Table) を使う準備。GetRunningObjectTable / CreateBindCtx を P/Invoke する。
$script:RotTypeReady = $false
function Initialize-RotType {
  if ($script:RotTypeReady) { return }
  Add-Type -Namespace PerfectVba -Name Ole32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("ole32.dll")]
public static extern int GetRunningObjectTable(int reserved, out System.Runtime.InteropServices.ComTypes.IRunningObjectTable prot);
[System.Runtime.InteropServices.DllImport("ole32.dll")]
public static extern int CreateBindCtx(int reserved, out System.Runtime.InteropServices.ComTypes.IBindCtx ppbc);
'@
  $script:RotTypeReady = $true
}

# 指定パスのブックを“今まさに開いている”インスタンスを ROT から直接取得する。
# GetActiveObject は起動中の先頭 1 インスタンスしか拾えないため、Excel が複数起動して
# いると対象ブックを取りこぼし、別インスタンスで読み取り専用コピーを開いて書き戻しが
# 闇に葬られる。ここでは ROT 参照のみで（＝Excel を起動せず）既に開かれているブックを返す。
function Get-RunningWorkbookByPath($full) {
  try { Initialize-RotType } catch { return $null }
  $rot = $null; $bc = $null; $enum = $null; $found = $null
  try {
    if ([PerfectVba.Ole32]::GetRunningObjectTable(0, [ref]$rot) -ne 0) { return $null }
    if ([PerfectVba.Ole32]::CreateBindCtx(0, [ref]$bc) -ne 0) { return $null }
    $rot.EnumRunning([ref]$enum)
    $mon = New-Object 'System.Runtime.InteropServices.ComTypes.IMoniker[]' 1
    while ($enum.Next(1, $mon, [System.IntPtr]::Zero) -eq 0) {
      $name = $null
      try { $mon[0].GetDisplayName($bc, $null, [ref]$name) } catch { $name = $null }
      # Excel のブックは ROT 表示名がフルパス。パス一致時のみ実体を取り出して二重確認する。
      if ($name -and ($name -ieq $full)) {
        $obj = $null
        try { [void]$rot.GetObject($mon[0], [ref]$obj) } catch { $obj = $null }
        if ($obj) {
          # パス一致で確定。URL 化(OneDrive 等)で FullName が食い違う場合は所有アプリで判定。
          $ok = $false
          try { $ok = ($obj.FullName -ieq $full) } catch { $ok = $false }
          if (-not $ok) { try { $ok = ([string]$obj.Application.Name -like '*Excel*') } catch { $ok = $false } }
          if ($ok) { $found = $obj }
        }
      }
      Release-Com $mon[0]
      if ($found) { break }
    }
  } catch {
    $found = $null
  } finally {
    Release-Com $enum
    Release-Com $bc
    Release-Com $rot
  }
  return $found
}

function New-ExcelContext($path, $headless, $attachOnly = $false) {
  $full = Resolve-OfficePath $path

  # 1) 対象ブックを開いている“実体のインスタンス”を ROT から直接掴む（複数起動対応）。
  $wb = Get-RunningWorkbookByPath $full
  $app = $null
  if ($wb) { try { $app = $wb.Application } catch { $app = $null } }

  # 2) ROT で取れなければ GetActiveObject + Workbooks 列挙にフォールバック。
  if (-not $wb) {
    try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $app = $null }
    if ($app) {
      foreach ($w in $app.Workbooks) {
        try { if ($w.FullName -ieq $full) { $wb = $w; break } } catch {}
      }
    }
  }

  # 3) 既に開かれている → ユーザーのインスタンスにそのまま接続（起動も再オープンもしない）。
  if ($wb) {
    return [pscustomobject]@{ App = $app; Doc = $wb; Started = $false; OpenedDoc = $false; Full = $full; OrigEvents = $null; NotOpen = $false }
  }

  # 4) attachOnly: 対象ブックが開かれていなければ Office を起動せず「未起動」を通知する。
  if ($attachOnly) {
    return [pscustomobject]@{ App = $app; Doc = $null; Started = $false; OpenedDoc = $false; Full = $full; OrigEvents = $null; NotOpen = $true }
  }

  # 5) 開かれていない → 自分で開く。
  $started = $false; $origEvents = $null
  if (-not $app) {
    $app = New-Object -ComObject Excel.Application
    $started = $true
    try { $app.Visible = (-not $headless) } catch {}
  }
  else {
    # 既存インスタンスを使う場合のみ、イベント設定を退避して後で戻す。
    try { $origEvents = $app.EnableEvents } catch {}
  }
  try { $app.DisplayAlerts = $false } catch {}
  try { $app.EnableEvents = $false } catch {}
  try { $app.AutomationSecurity = 3 } catch {}  # msoAutomationSecurityForceDisable: マクロ自動実行を抑止
  if (-not (Test-Path $full)) { throw "ファイルが見つかりません: $full" }
  $wb = $app.Workbooks.Open($full)
  return [pscustomobject]@{ App = $app; Doc = $wb; Started = $started; OpenedDoc = $true; Full = $full; OrigEvents = $origEvents; NotOpen = $false }
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

function New-AccessContext($path, $headless, $attachOnly = $false) {
  $full = Resolve-OfficePath $path
  $started = $false; $openedDoc = $false
  try { $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Access.Application') } catch { $app = $null }

  if ($app) {
    $cur = $null
    try { $cur = $app.CurrentProject.FullName } catch { $cur = $null }
    if ($cur -and ($cur -ieq $full)) {
      # ユーザーが開いている当該 DB にそのまま接続
      return [pscustomobject]@{ App = $app; Started = $false; OpenedDoc = $false; Full = $full; NotOpen = $false }
    }
    # 別の DB を開いている既存インスタンスは使わず、新規インスタンスを起動する
    Release-Com $app
    $app = $null
  }

  # attachOnly: 対象 DB が開かれていなければ Office を起動せず「未起動」を通知する
  if ($attachOnly) {
    return [pscustomobject]@{ App = $null; Started = $false; OpenedDoc = $false; Full = $full; NotOpen = $true }
  }

  $app = New-Object -ComObject Access.Application
  $started = $true
  try { $app.Visible = (-not $headless) } catch {}
  if (-not (Test-Path $full)) { throw "ファイルが見つかりません: $full" }
  try { $app.AutomationSecurity = 3 } catch {}
  # 排他オープン（SaveAsText/LoadFromText に必要）
  $app.OpenCurrentDatabase($full, $true)
  $openedDoc = $true
  return [pscustomobject]@{ App = $app; Started = $started; OpenedDoc = $openedDoc; Full = $full; NotOpen = $false }
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

# VBA モジュール名の比較用正規化。識別子は大文字小文字を区別しないため小文字化し、
# 前後空白とユニコード合成の揺れ（濁点など）を吸収する。全角/半角は別名のため変えない。
function Normalize-CompName($name) {
  if ($null -eq $name) { return '' }
  $s = [string]$name
  try { $s = $s.Normalize([System.Text.NormalizationForm]::FormC) } catch {}
  return $s.Trim().ToLowerInvariant()
}

# 名前一致するコンポーネントを返す。完全一致を最優先し、無ければ正規化一致でフォールバック。
# 正規化一致を見ることで、ファイル名由来の前後空白・大小・合成の揺れで既存モジュールを
# 取りこぼして無用な新規追加（→ 重複モジュール）を作る事故を防ぐ。
function Find-Component($proj, $name) {
  $norm = Normalize-CompName $name
  $fallback = $null
  foreach ($c in $proj.VBComponents) {
    $cn = $c.Name
    if ($cn -eq $name) { return $c }
    if (-not $fallback -and (Normalize-CompName $cn) -eq $norm) { $fallback = $c }
  }
  return $fallback
}

# std/class モジュールを新規追加し、確実に目的名を付ける。
# 名前衝突等でリネームできない場合は、作りかけのゴミモジュールを残さず明確なエラーにする。
# （従来は Name 設定失敗を握りつぶし、Module1 等のゴミに編集が入って重複・未反映を招いていた）
function Add-StdOrClassComponent($proj, $name, $kind) {
  $type = if ($kind -eq 'class') { $script:VBEXT_CLASS } else { $script:VBEXT_STD }
  $comp = $proj.VBComponents.Add($type)
  try {
    $comp.Name = $name
  }
  catch {
    try { $proj.VBComponents.Remove($comp) } catch {}
    throw "モジュール '$name' を作成できません。同じ名前のモジュール/クラス/オブジェクトが既に存在するか、名前に使えない文字（前後の空白・記号・全角半角の揺れ等）が含まれています。VBE 側で重複したモジュールを削除し、ファイル名を正しいモジュール名に直してから再実行してください。"
  }
  return $comp
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
