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
      # std/class: 既存・新規とも CodeModule 直接操作で反映する（Excel と同じ方式）。
      # Import(Remove+Import) は VBE のコードウィンドウを開いたままだと no-op 化したり、
      # Remove の遅延で重複モジュール(Module1/Module11 等)を生み編集が反映されないため使わない。
      $target = Find-Component $proj $comp.name
      # フォーム/レポートのコードビハインド(Form_xxx 等, Type=100)と名前が衝突していると、
      # 黙ってスキップ（未反映）になっていた。明確にエラーで知らせる。
      if ($target -and $target.Type -eq $script:VBEXT_DOCUMENT) {
        throw "モジュール '$($comp.name)' は Access のフォーム/レポートのコードビハインド（$($target.Name)）と名前が衝突しています。標準/クラスモジュールには別の名前を付けてください。"
      }
      if (-not $target) {
        # 新規追加。名前を付けられない場合はゴミを残さず明確なエラーにする（旧: 握りつぶし）。
        $target = Add-StdOrClassComponent $proj $comp.name $comp.kind
      }
      if ($null -ne $comp.codeText) {
        Set-ModuleCodeInPlace $target $comp.codeText
      }
      # 編集をモジュールとして DB に保存する。ユーザーが既に開いているインスタンスへ
      # attach した場合はクローズでコミットされないため、ここで明示的に保存する。
      # リネーム後に実際に付いた名前($target.Name)で保存する（目的名で未保存になる事故を防ぐ）。
      try { $ctx.App.DoCmd.Save($script:AC_MODULE, $target.Name) } catch {}
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
