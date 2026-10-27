<#
  QR Hub(MKQR) SharePoint 목록 · 열 · 인덱스 한 번에 맞추기 (2026-10-27) — 빠진 것만 만든다
  ---------------------------------------------------------------------------------------------
  지금 코드가 쓰는 MKQR 목록 전부(STK-DB 사이트)를 점검해 없는 목록 · 열은 만들고, 필요한 인덱스를 건다.
  이미 있는 목록 · 열은 그대로 둔다(형식도 바꾸지 않는다) — 다시 실행해도 안전.
  끝에 목록마다 「빠진 열 · 형식이 다른 열 · 인덱스」 를 표로 보여 준다.

    1) mkqr_Centers          센터(방문 인증)
    2) mkqr_Agents           Agent(딜러사 명함 QR)
    3) mkqr_CheckIns         센터 · Agent 체크인 기록         인덱스 ServerTimestamp · CenterId · DeviceFingerprint
    4) mkqr_CategoryAccess   화면별 권한(Access_* 예/아니요 · ViewOnly)
    5) mkqr_AccessLog        접속 기록                        인덱스 LoginAt
    6) mkqr_ITAssets         IT 자산                          인덱스 QrToken · CICD · SerialNo · AssigneeEmail · Assignee
    7) mkqr_ITScans          IT 자산 스캔(ScanType 포함)      인덱스 Created · AssetId · ScannedBy
    8) mkqr_ITStock          액세서리 재고
    9) mkqr_ITHistory        지급 · 반납 · 교체 이력           인덱스 PersonEmail · AssetId · ChangeDate · ImportKey
   10) mkqr_ITSecrets        BitLocker 복구 키(IT 전용)        인덱스 SerialNo · AssetId   ※ 권한 상속 중지 · IT 담당만
   11) mkqr_ITDeleteRequests 자산 삭제 결재                    인덱스 ApproverEmail · AssetId · Status

  ※ 새로 만드는 Access_* 열은 기본값을 비워 둔다 — 비어 있으면 「기존 규칙(허용 · 관리자 따르기)」 이 적용되고,
    기본값이 「아니요」 면 새로 추가하는 사용자가 그 화면에서 막히기 때문.
  ※ 목록 하나가 5,000건을 넘었으면 인덱스 만들기가 실패할 수 있다 — 그 경우 표에 「인덱스 실패」 로 나온다.

  사용 (mkqr-it-lists.ps1 과 같은 방식 · 일반 모드 ISE):
    .\mkqr-all-lists.ps1 -ClientId <지난번 열 만들기 때 쓴 PnP 클라이언트 ID>
    .\mkqr-all-lists.ps1 -ClientId <ID> -CheckOnly      # 만들지 않고 점검 표만
  ※ 시크릿 · 토큰을 이 파일에 적지 않는다.
#>
param(
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$SiteUrl = "https://startruckkorea.sharepoint.com/sites/STK-DB",
  [switch]$CheckOnly
)
$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

# ── 목록 정의 ─────────────────────────────────────────────────────────────────────────────
# T = 한 줄 텍스트, N = 숫자, B = 예/아니요, M = 여러 줄 텍스트
$ACCESS = "Access_overview", "Access_centers", "Access_qr", "Access_checkins", "Access_agents", "Access_agentQr", "Access_agentCheckins",
          "Access_itAssets", "Access_itQr", "Access_itScans", "Access_admins", "Access_accessLog", "Access_logs"
$Spec = [ordered]@{
  "mkqr_Centers" = @{
    T = "Region", "CenterNameEn", "CorpNameKr", "AddressKr", "AddressEn", "Phone", "RedirectUrl", "QrToken", "IconDriveItemId"
    N = "Lat", "Lng", "VerifyRadiusM"; B = @("Active"); M = @(); I = @(); BDefault = @{ Active = "1" } }
  "mkqr_Agents" = @{
    T = "AgentNameEn", "CorpNameKr", "AddressKr", "AddressEn", "Phone", "RedirectUrl", "QrToken"
    N = @(); B = @("Active"); M = @(); I = @(); BDefault = @{ Active = "1" } }
  "mkqr_CheckIns" = @{
    T = "CenterId", "CenterName", "ServerTimestamp", "ClientTimestamp", "PublicIP", "IpCity", "DistanceSource", "VerifyStatus", "UserAgent", "DeviceId", "DeviceFingerprint"
    N = "IpLat", "IpLng", "GpsLat", "GpsLng", "GpsAccuracy", "DistanceKm", "RecaptchaScore"; B = @(); M = @()
    I = "ServerTimestamp", "CenterId", "DeviceFingerprint" }
  "mkqr_CategoryAccess" = @{
    T = "Email", "ViewOnly"; N = @(); B = $ACCESS; M = @(); I = @("Email"); BDefault = @{} }
  "mkqr_AccessLog" = @{
    T = "UserEmail", "UserName", "Role", "LoginAt", "LastActivityAt", "LogoutAt", "EndType", "Device"; N = @(); B = @(); M = @("UserAgent")
    I = @("LoginAt") }
  "mkqr_ITAssets" = @{
    T = "SerialNo", "AssetCode", "CICD", "QrToken", "AssetType", "Location", "AddressKr", "Assignee", "AssigneeEmail", "RentalEndDate", "RentalVendor",
        "State", "Model", "Hostname", "AssignedDate", "PhoneNo", "CriticalApps", "SecurityApp", "ExtCiId", "ExtObjectId", "WarrantyStart", "SourceKey"
    N = "Lat", "Lng", "VerifyRadiusM"; B = "IsRental", "Active"; M = @("Note")
    I = "QrToken", "CICD", "SerialNo", "AssigneeEmail", "Assignee"; BDefault = @{ Active = "1"; IsRental = "0" } }
  "mkqr_ITScans" = @{
    T = "AssetId", "CICD", "AssetTitle", "ScannedBy", "ScannedByName", "ClientTimestamp", "VerifyStatus", "UserAgent", "ScanType"
    N = "GpsLat", "GpsLng", "GpsAccuracy", "DistanceM"; B = @(); M = @(); I = "Created", "AssetId", "ScannedBy" }
  "mkqr_ITStock" = @{
    T = "Category", "Entity", "Ownership"; N = "Qty", "PoolQty", "Threshold"; B = "IsAsset", "Active"; M = @("Note"); I = @()
    BDefault = @{ Active = "1" } }
  "mkqr_ITHistory" = @{
    T = "ChangeDate", "Kind", "AssetId", "CICD", "StockId", "StockKey", "ItemType", "Model", "SerialNo", "ChangeType", "Person", "PersonEmail", "Dept",
        "PrevState", "NewState", "Reason", "RequestedBy", "ApprovedBy", "HandledBy", "TicketNo", "Source", "ImportKey"
    N = @("Qty"); B = @(); M = @("Note"); I = "PersonEmail", "AssetId", "ChangeDate", "ImportKey" }
  "mkqr_ITSecrets" = @{
    T = "SerialNo", "AssetId", "Hostname", "Model", "Owner", "Source"; N = @(); B = @(); M = "RecoveryKey", "Note"; I = "SerialNo", "AssetId"; Secret = $true }
  "mkqr_ITDeleteRequests" = @{
    T = "AssetId", "CICD", "AssetTitle", "SerialNo", "RequesterEmail", "RequesterName", "ApproverEmail", "ApproverName", "Status", "DecidedAt", "DecidedBy", "DeletedAt"
    N = @(); B = @(); M = "Reason", "DecisionNote", "Snapshot"; I = "ApproverEmail", "AssetId", "Status" }
}
$TypeOf = @{ T = "Text"; N = "Number"; B = "Boolean"; M = "Note" }

# ── 만들기 ──────────────────────────────────────────────────────────────────────────────
$made = New-Object System.Collections.Generic.List[string]
$idxFail = @{}
foreach ($name in $Spec.Keys) {
  $s = $Spec[$name]
  Write-Host "`n[$name]" -ForegroundColor Cyan
  $list = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
  if (-not $list) {
    if ($CheckOnly) { Write-Host "  목록 없음" -ForegroundColor Yellow; continue }
    $list = New-PnPList -Title $name -Template GenericList -OnQuickLaunch:$false
    $made.Add("$name (목록)"); Write-Host "  목록 생성" -ForegroundColor Green
  }
  if ($CheckOnly) { continue }
  $have = @{}; Get-PnPField -List $name | ForEach-Object { $have[$_.InternalName] = $_ }
  foreach ($k in "T", "N", "B", "M") {
    foreach ($col in @($s[$k])) {
      if (-not $col -or $have.ContainsKey($col)) { continue }
      Add-PnPField -List $name -DisplayName $col -InternalName $col -Type $TypeOf[$k] -AddToDefaultView | Out-Null
      if ($k -eq "B") {
        $def = if ($s.BDefault -and $s.BDefault.ContainsKey($col)) { $s.BDefault[$col] } else { "" }
        Set-PnPField -List $name -Identity $col -Values @{ DefaultValue = $def } | Out-Null
      }
      $made.Add("$name . $col ($($TypeOf[$k]))"); Write-Host "  열 생성: $col ($($TypeOf[$k]))" -ForegroundColor Green
    }
  }
  foreach ($col in @($s.I)) {
    if (-not $col) { continue }
    try {
      $f = Get-PnPField -List $name -Identity $col
      if (-not $f.Indexed) { Set-PnPField -List $name -Identity $col -Values @{ Indexed = $true } | Out-Null; $made.Add("$name . $col (인덱스)"); Write-Host "  인덱스: $col" -ForegroundColor Green }
    } catch { $idxFail["$name|$col"] = $_.Exception.Message; Write-Host "  인덱스 실패: $col — $($_.Exception.Message)" -ForegroundColor Red }
  }
  if ($s.Secret) { Write-Host "  !! 목록 설정 > 이 목록에 대한 권한 > 권한 상속 중지 > 구성원 · 방문자 제거 > IT 담당자만 (상속이 이어지면 전 직원이 키를 읽음)" -ForegroundColor Yellow }
}

# ── 점검 표 ─────────────────────────────────────────────────────────────────────────────
Write-Host "`n════════ 점검 결과 ════════" -ForegroundColor Cyan
$problems = 0
foreach ($name in $Spec.Keys) {
  $s = $Spec[$name]
  $list = Get-PnPList -Identity $name -ErrorAction SilentlyContinue
  if (-not $list) { Write-Host ("  {0,-24} 목록 없음" -f $name) -ForegroundColor Red; $problems++; continue }
  $have = @{}; Get-PnPField -List $name | ForEach-Object { $have[$_.InternalName] = $_ }
  $miss = @(); $diff = @(); $noIdx = @()
  foreach ($k in "T", "N", "B", "M") {
    foreach ($col in @($s[$k])) {
      if (-not $col) { continue }
      if (-not $have.ContainsKey($col)) { $miss += $col; continue }
      $t = $have[$col].TypeAsString; $want = $TypeOf[$k]
      $ok = ($t -eq $want) -or ($want -eq "Text" -and $t -in "Note", "DateTime") -or ($want -eq "Number" -and $t -in "Integer", "Currency") -or ($want -eq "Note" -and $t -eq "Text")
      if (-not $ok) { $diff += "$col($t→$want)" }
    }
  }
  foreach ($col in @($s.I)) { if ($col -and $have.ContainsKey($col) -and -not $have[$col].Indexed) { $noIdx += $col } }
  $line = "  {0,-24} {1,6:N0}건" -f $name, $list.ItemCount
  if (-not $miss -and -not $noIdx -and -not $diff) { Write-Host "$line  정상" -ForegroundColor Green; continue }
  Write-Host $line
  if ($miss)  { Write-Host ("      빠진 열: " + ($miss -join ", ")) -ForegroundColor Red; $problems++ }
  if ($noIdx) { Write-Host ("      인덱스 없음: " + ($noIdx -join ", ") + $(if ($list.ItemCount -ge 5000) { "  (5,000건 이상 — 인덱스 생성이 막힐 수 있음)" } else { "" })) -ForegroundColor Yellow; $problems++ }
  if ($diff)  { Write-Host ("      형식 다름(참고): " + ($diff -join ", ")) -ForegroundColor DarkYellow }
}
Write-Host ""
if ($made.Count) { Write-Host ("이번에 만든 것 {0}개:" -f $made.Count) -ForegroundColor Green; $made | ForEach-Object { Write-Host "  + $_" } }
else { Write-Host "새로 만든 것 없음." }
if ($problems) { Write-Host "`n확인이 필요한 항목이 있습니다(위 빨강 · 노랑)." -ForegroundColor Yellow }
else { Write-Host "`n모든 MKQR 목록 · 열 · 인덱스가 갖춰져 있습니다. QR Hub 를 새로고침하세요." -ForegroundColor Green }
