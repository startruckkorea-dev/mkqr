<#
  MKQR IT 자산 QR · CMDB(10-23 · mkqr-it · mkqr-cmdb) — SharePoint 목록 · 열 준비 (한 번만 실행)
  ---------------------------------------------------------------------------
  만드는 것 (STK-DB 사이트):
    1) mkqr_ITAssets  — IT 자산 (Description=제목, Serial No., Asset Code, CI CD, QR 토큰, 위치 좌표 · 허용 반경 등)
    2) mkqr_ITScans   — QR 스캔 기록 (스캔한 사람 · GPS · 거리 · 판정)
    3) mkqr_CategoryAccess 에 권한 열 3개 — Access_itAssets · Access_itQr · Access_itScans (예/아니요)
    4) mkqr_ITStock   — 수량으로 관리하는 액세서리 재고 (신품 · 회수 재고 · 최소 기준)
    5) mkqr_ITHistory — 지급 · 반납 · 교체 · 고장 · 폐기 이력 (자산 · 액세서리 공용)
    6) mkqr_ITSecrets — BitLocker 복구 키(IT 전용)
    7) mkqr_ITDeleteRequests — 자산 삭제 결재 요청(10-27)
  ※ 목록을 새로 만들면 사이트 권한을 그대로 이어받는다(권한 상속을 끊지 않는다). 상속을 끊으면 IT 담당자에게 「편집」 을 따로 줘야 한다.
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
foreach ($n in "SerialNo", "AssetCode", "CICD", "QrToken", "AssetType", "Location", "AddressKr", "Assignee", "AssigneeEmail", "RentalEndDate", "RentalVendor") { Ensure-Field $A $n "Text" }   # Assignee = 이름, AssigneeEmail = 메일(직원 명부 또는 직접 입력)
# CMDB(10-23) — 상태 · 모델 · 호스트 · 지급일 · 전화번호 · 프로그램 · DaaS CI · 보증 · 가져오기 키(날짜도 한국 날짜 YYYY-MM-DD 글자)
foreach ($n in "State", "Model", "Hostname", "AssignedDate", "PhoneNo", "CriticalApps", "SecurityApp", "ExtCiId", "ExtObjectId", "WarrantyStart", "SourceKey") { Ensure-Field $A $n "Text" }
Ensure-Field $A "IsRental" "Boolean"   # 렌탈 장비 여부 — 종료일은 한국 날짜 YYYY-MM-DD 글자로 저장
foreach ($n in "Lat", "Lng", "VerifyRadiusM") { Ensure-Field $A $n "Number" }
Ensure-Field $A "Note" "Note"
Ensure-Field $A "Active" "Boolean"
Set-PnPField -List $A -Identity "Active" -Values @{ DefaultValue = "1" } | Out-Null
Ensure-Index $A "QrToken"; Ensure-Index $A "CICD"; Ensure-Index $A "SerialNo"; Ensure-Index $A "AssigneeEmail"

# 2) 스캔 기록
$S = "mkqr_ITScans"; Ensure-List $S | Out-Null
foreach ($n in "AssetId", "CICD", "AssetTitle", "ScannedBy", "ScannedByName", "ClientTimestamp", "VerifyStatus", "UserAgent", "ScanType") { Ensure-Field $S $n "Text" }   # ScanType(10-27) = 본인 확인 · 본인 확인(PC) · IT 점검 · 습득 · 타인
foreach ($n in "GpsLat", "GpsLng", "GpsAccuracy", "DistanceM") { Ensure-Field $S $n "Number" }
Ensure-Index $S "AssetId"; Ensure-Index $S "Created"; Ensure-Index $S "ScannedBy"   # ScannedBy(10-27) — 내 자산의 본인 확인 표시

# 3) 권한 열 — 비어 있으면 관리자(Access_admins) 값을 따르고, 그것도 없으면 「없음」
$C = "mkqr_CategoryAccess"
foreach ($n in "Access_itAssets", "Access_itQr", "Access_itScans") { Ensure-Field $C $n "Boolean" }

# 4) 재고(액세서리)
$K = "mkqr_ITStock"; Ensure-List $K | Out-Null
foreach ($n in "Category", "Entity", "Ownership") { Ensure-Field $K $n "Text" }
foreach ($n in "Qty", "PoolQty", "Threshold") { Ensure-Field $K $n "Number" }
foreach ($n in "IsAsset", "Active") { Ensure-Field $K $n "Boolean" }
Set-PnPField -List $K -Identity "Active" -Values @{ DefaultValue = "1" } | Out-Null
Ensure-Field $K "Note" "Note"

# 5) 변경 이력 — 사람(PersonEmail) · 자산(AssetId) · 날짜로 찾는다
$H = "mkqr_ITHistory"; Ensure-List $H | Out-Null
foreach ($n in "ChangeDate", "Kind", "AssetId", "CICD", "StockId", "StockKey", "ItemType", "Model", "SerialNo", "ChangeType", "Person", "PersonEmail", "Dept",
  "PrevState", "NewState", "Reason", "RequestedBy", "ApprovedBy", "HandledBy", "TicketNo", "Source", "ImportKey") { Ensure-Field $H $n "Text" }
Ensure-Field $H "Qty" "Number"
Ensure-Field $H "Note" "Note"
Ensure-Index $H "PersonEmail"; Ensure-Index $H "AssetId"; Ensure-Index $H "ChangeDate"; Ensure-Index $H "ImportKey"

# 6) BitLocker 복구 키 — IT 전용 목록. 만든 뒤 반드시 「권한 상속 중지」 → 구성원 · 방문자 제거 → IT 담당자만 남긴다
$X = "mkqr_ITSecrets"; Ensure-List $X | Out-Null
foreach ($n in "SerialNo", "AssetId", "Hostname", "Model", "Owner", "Source") { Ensure-Field $X $n "Text" }
foreach ($n in "RecoveryKey", "Note") { Ensure-Field $X $n "Note" }
Ensure-Index $X "SerialNo"; Ensure-Index $X "AssetId"
Write-Host "  !! $X : 목록 설정 > 이 목록에 대한 권한 > 권한 상속 중지 > 구성원 · 방문자 제거 > IT 담당자만 (전 직원 참가가 이어지면 누구나 키를 읽음)" -ForegroundColor Yellow

# 7) 자산 삭제 결재(10-27) — 삭제 요청 · 결재 기록. 결재자 = 요청자의 직원 명부 Boss
$R = "mkqr_ITDeleteRequests"; Ensure-List $R | Out-Null
foreach ($n in "AssetId", "CICD", "AssetTitle", "SerialNo", "RequesterEmail", "RequesterName", "ApproverEmail", "ApproverName", "Status", "DecidedAt", "DecidedBy", "DeletedAt") { Ensure-Field $R $n "Text" }
foreach ($n in "Reason", "DecisionNote", "Snapshot") { Ensure-Field $R $n "Note" }
Ensure-Index $R "ApproverEmail"; Ensure-Index $R "AssetId"; Ensure-Index $R "Status"

Write-Host "`n완료 — MKQR 를 새로고침한 뒤 「카테고리 접근 권한」 에서 IT 담당자에게 IT 전용 칸을 체크하고 「반영」 하세요."
