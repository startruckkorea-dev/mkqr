# MKQR Worker 배포 — GitHub 의 최신 worker.js 를 받아 이 PC 의 Worker 폴더 설정(wrangler.toml · 비밀값)으로 배포한다.
# 실행(PowerShell 한 줄):
#   irm https://raw.githubusercontent.com/startruckkorea-dev/mkqr/main/worker/deploy.ps1 | iex
# 하는 일: ① 폴더 · 이름 확인(다른 Worker 덮어쓰기 방지) ② 최신 코드 받기 ③ 기존 파일 백업 ④ wrangler deploy ⑤ 배포된 버전 확인
# 배포가 실패하면 백업으로 되돌린다. 비밀값(MK_CLIENT_SECRET 등)은 Cloudflare 에 그대로 남아 있어 다시 넣을 필요 없음.
& {
  $ErrorActionPreference = 'Stop'
  $dir  = if ($env:MKQR_WORKER_DIR) { $env:MKQR_WORKER_DIR } else { 'C:\mkqr\worker' }
  $raw  = if ($env:MKQR_RAW) { $env:MKQR_RAW } else { 'https://raw.githubusercontent.com/startruckkorea-dev/mkqr/main/worker/worker.js' }
  $api  = if ($env:MKQR_API) { $env:MKQR_API } else { 'https://mkqr-worker.mkqr.workers.dev/api/version' }
  $tmpDir = if ($env:TEMP) { $env:TEMP } else { [IO.Path]::GetTempPath() }
  function Say($m, $c = 'Gray') { Write-Host $m -ForegroundColor $c }

  Say "MKQR Worker 배포 — 폴더 $dir" 'Cyan'
  $toml = Join-Path $dir 'wrangler.toml'
  if (-not (Test-Path $toml)) { Say "중단: $toml 이 없습니다. Worker 폴더를 확인하세요(다른 폴더면 `$env:MKQR_WORKER_DIR='경로' 를 먼저 지정)." 'Red'; return }
  $t = Get-Content $toml -Raw
  $top = ($t -split '(?m)^\s*\[')[0]   # 첫 [구역] 앞(최상위)만 본다 — [env.x] 안의 name 에 속지 않게
  if ($top -notmatch '(?m)^\s*name\s*=\s*"mkqr-worker"\s*$') { Say '중단: wrangler.toml 의 name 이 "mkqr-worker" 가 아닙니다 — 다른 Worker 에 덮어쓰지 않도록 멈춥니다.' 'Red'; return }
  $main = if ($top -match '(?m)^\s*main\s*=\s*"([^"]+)"') { $Matches[1] } else { 'worker.js' }
  $target = Join-Path $dir $main

  Say '① 최신 코드 받는 중...'
  $tmp = Join-Path $tmpDir ("mkqr-worker-{0}.js" -f (Get-Date -Format 'yyyyMMddHHmmss'))
  try { Invoke-WebRequest -Uri ($raw + '?t=' + [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()) -OutFile $tmp -UseBasicParsing }
  catch { Say "중단: GitHub 에서 받지 못했습니다 — $($_.Exception.Message)" 'Red'; return }
  $code = [IO.File]::ReadAllText($tmp, [Text.Encoding]::UTF8)
  if ($code -notmatch 'export default' -or $code -notmatch 'WORKER_VERSION = "([^"]+)"') { Say '중단: 받은 파일이 Worker 코드가 아닙니다.' 'Red'; return }
  $want = $Matches[1]
  Say "   받은 버전: $want"

  $bak = $null
  if (Test-Path $target) {
    if ((Get-FileHash $target).Hash -eq (Get-FileHash $tmp).Hash) { Say '   지금 폴더의 코드와 같습니다 — 그래도 다시 배포합니다.' }
    $bak = "$target.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Copy-Item $target $bak
    Say "② 기존 파일 백업: $bak"
  }
  Copy-Item $tmp $target -Force

  Say '③ wrangler deploy ...'
  Push-Location $dir
  try { & wrangler deploy; $ok = ($LASTEXITCODE -eq 0) } catch { $ok = $false; Say $_.Exception.Message 'Red' } finally { Pop-Location }
  if (-not $ok) {
    if ($bak) { Copy-Item $bak $target -Force; Say '배포 실패 — 폴더의 파일을 백업으로 되돌렸습니다(실제 Worker 는 바뀌지 않음).' 'Red' }
    else { Remove-Item $target -Force; Say '배포 실패 — 받은 파일을 지웠습니다(실제 Worker 는 바뀌지 않음).' 'Red' }
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    return
  }

  Remove-Item $tmp -Force -ErrorAction SilentlyContinue
  Say '④ 배포된 버전 확인 ...'
  Start-Sleep -Seconds 3
  try {
    $v = Invoke-RestMethod -Uri ($api + '?t=' + [DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
    if ($v.version -eq $want) { Say "완료: mkqr-worker $($v.version) · 보관 연도 $($v.years -join ', ') · 지금 기록 연도 $($v.writeYear)" 'Green' }
    else { Say "확인 필요: 배포된 버전 $($v.version) / 기대 $want — 1~2분 뒤 $api 를 다시 열어 보세요." 'Yellow' }
  } catch { Say "확인 필요: $api 응답 없음 — $($_.Exception.Message)" 'Yellow' }
}
