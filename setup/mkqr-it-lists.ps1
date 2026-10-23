<#
  MKQR IT 자산 QR(10-23 · mkqr-it) — SharePoint 목록 · 열 준비 (한 번만 실행)
  ---------------------------------------------------------------------------
  만드는 것 (STK-DB 사이트):
    1) mkqr_ITAssets  — IT 자산 (Description=제목, Serial No., Asset Code, CI CD, QR 토큰, 위치 좌표 · 허용 반경 등)
    2) mkqr_ITScans   — QR 스캔 기록 (스캔한 사람 · GPS · 거리 · 판정)
    3) mkqr_CategoryAccess 에 권한 열 3개 — Access_itAssets · Access_itQr · Access_itScans (예/아니요)
  이미 있는 목록 · 열은 건너뛴다(다시 실행해도 안전).

  사용:
    Install-Module PnP.PowerShell -Scope CurrentUser      # 처음 한 번
    .\mkqr-it-lists.ps1 -ClientId <PnP 용 Entra 앱 클라이언트 ID>
  ※ 시크릿 · 토큰을 이 파일에 적지 않는다. 로그인은 -Interactive(본인 계정 · 사이트 소유자 권한).
#>
param(
  [Parameter(Mandatory = $true)][string]$ClientId,
  [string]$SiteUrl = "https://startruckkorea.sharepoint.com/sites/STK-DB"
)
$ErrorActionPreference = "Stop"
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive

function Ensure-List([string]$title) {
  $l = Get-PnPList -Identity $title -ErrorAction SilentlyContinue
  if (-not $l) { $l = New-PnPList -Title $title -Template GenericList -OnQuickLaunch:$false; Write-Host "목록 생성: $title" }
  else { Write-Host "목록 있음: $title" }
  return $l
}
function Ensure-Field([string]$list, [string]$name, [string]$type) {
  $f = Get-PnPField -List $list -Identity $name -ErrorAction SilentlyContinue
  if ($f) { Write-Host "  열 있음: $name"; return }
  Add-PnPField -List $list -DisplayName $name -InternalName $name -Type $type -AddToDefaultView | Out-Null
  Write-Host "  열 생성: $name ($type)"
}
function Ensure-Index([string]$list, [string]$name) {
  Set-PnPField -List $list -Identity $name -Values @{ Indexed = $true } | Out-Null
  Write-Host "  인덱스: $name"
}

# 1) IT 자산
$A = "mkqr_ITAssets"; Ensure-List $A | Out-Null
foreach ($n in "SerialNo", "AssetCode", "CICD", "QrToken", "AssetType", "Location", "AddressKr", "Assignee", "RentalEndDate", "RentalVendor") { Ensure-Field $A $n "Text" }
Ensure-Field $A "IsRental" "Boolean"   # 렌탈 장비 여부 — 종료일은 한국 날짜 YYYY-MM-DD 글자로 저장
foreach ($n in "Lat", "Lng", "VerifyRadiusM") { Ensure-Field $A $n "Number" }
Ensure-Field $A "Note" "Note"
Ensure-Field $A "Active" "Boolean"
Set-PnPField -List $A -Identity "Active" -Values @{ DefaultValue = "1" } | Out-Null
Ensure-Index $A "QrToken"; Ensure-Index $A "CICD"

# 2) 스캔 기록
$S = "mkqr_ITScans"; Ensure-List $S | Out-Null
foreach ($n in "AssetId", "CICD", "AssetTitle", "ScannedBy", "ScannedByName", "ClientTimestamp", "VerifyStatus", "UserAgent") { Ensure-Field $S $n "Text" }
foreach ($n in "GpsLat", "GpsLng", "GpsAccuracy", "DistanceM") { Ensure-Field $S $n "Number" }
Ensure-Index $S "AssetId"; Ensure-Index $S "Created"

# 3) 권한 열 — 비어 있으면 관리자(Access_admins) 값을 따르고, 그것도 없으면 「없음」
$C = "mkqr_CategoryAccess"
foreach ($n in "Access_itAssets", "Access_itQr", "Access_itScans") { Ensure-Field $C $n "Boolean" }

Write-Host "`n완료 — MKQR 를 새로고침한 뒤 「카테고리 접근 권한」 에서 IT 담당자에게 IT 전용 칸을 체크하고 「반영」 하세요."
