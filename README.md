# MKQR — 캠페인/현장 방문 인증 QR 시스템

Star Truck Korea 서비스센터별 방문 인증용 QR 코드 시스템. ICS/APS와 동일한 아키텍처 패턴(React+Babel 단일 HTML, MSAL, SharePoint STK-DB)을 재사용하되, QR을 스캔하는 사람은 Hyosung 계정으로 로그인하지 않으므로 그 부분만 Cloudflare Worker(앱 전용 인증)로 분리했습니다.

## 구성

| 파일 | 역할 | 인증 |
|---|---|---|
| `site/index.html` | 관리자 화면 — 센터 관리 / QR 생성 / 체크인 기록 조회·승인 | MSAL 로그인 (Hyosung 계정) |
| `site/checkin.html` | QR 스캔 시 뜨는 공개 체크인 페이지 | 없음 (공개) |
| `worker/worker.js` | 체크인 페이지가 호출하는 API. 앱 자격증명으로 SharePoint에 대신 씀 | 앱 전용(client credentials) |

## 왜 Worker가 필요한가

기존 APS 시스템들은 전부 "로그인한 임직원"이 자기 권한으로 SharePoint에 씁니다. 하지만 mkqr은 방문자가 QR을 찍는 방식이라 로그인 주체가 없고, GitHub Pages는 정적 호스팅이라 비밀값(client secret)을 안전하게 보관할 서버가 없습니다. 그래서 체크인 저장 API만 별도 서버(Cloudflare Worker)에서 앱 전용 인증으로 처리합니다. 관리자 화면(index.html)은 기존 패턴 그대로 MSAL을 씁니다.

## 배포 순서

1. **SharePoint 리스트 3개 생성** (STK-DB 사이트, 아래 스키마)
2. **Entra ID 앱 등록** (Worker 전용, APS-MasterData-App과 분리 권장)
   - API 권한 → Application permissions → `Sites.Selected` → 관리자 동의
   - PnP PowerShell로 STK-DB 사이트에 이 앱의 쓰기 권한 부여:
     ```powershell
     Connect-PnPOnline -Url https://startruckkorea.sharepoint.com/sites/STK-DB -Interactive
     Grant-PnPAzureADAppSitePermission -AppId <새앱ID> -DisplayName "MKQR-Worker" -Site https://startruckkorea.sharepoint.com/sites/STK-DB -Permissions Write
     ```
   - 클라이언트 시크릿 생성 (만료일 캘린더 등록 권장)
3. **Cloudflare Worker 배포**
   ```bash
   cd worker
   wrangler secret put MK_CLIENT_ID
   wrangler secret put MK_CLIENT_SECRET
   wrangler secret put MK_TENANT_ID     # 19cab1f5-21f4-44df-8ac6-96d6ca595203
   wrangler deploy
   ```
4. **URL 반영**: 배포된 Worker 주소를 `site/index.html`의 `CHECKIN_API_BASE`, `site/checkin.html`의 `API_BASE`에 반영
5. **GitHub Pages 배포**: `site/` 내용을 `mkqr` 저장소 루트(main)에 push → GitHub Pages 자동 배포
6. **관리자 화면 접속** → 센터 관리에서 엑셀 일괄 가져오기(서비스센터직원관리대장 형식) → 각 센터 "좌표 찾기"로 위경도 채우기 → QR 코드 탭에서 다운로드/인쇄

## SharePoint 리스트 생성 — 화면 클릭 순서

STK-DB 사이트(`https://startruckkorea.sharepoint.com/sites/STK-DB`)에서 진행합니다.

**공통 절차**: 사이트 접속 → 왼쪽 메뉴 `사이트 콘텐츠(Site contents)` → 상단 `+ 새로 만들기(New)` → `목록(List)` → `빈 목록(Blank list)` → 이름 입력(아래 리스트명 그대로) → `만들기(Create)`.

목록이 만들어지면 기본적으로 **제목(Title)** 컬럼 하나만 있습니다. 나머지 컬럼은 목록 화면에서 맨 오른쪽 **`+ 열 추가(Add column)`**를 눌러 하나씩 추가합니다. 컬럼 추가 시 "열 이름"에는 반드시 **영문 컬럼명**(아래 표의 컬럼란, 대소문자까지 정확히)을 입력하세요 — 코드가 이 이름으로 값을 읽고 씁니다. "이 열 유형"에서 표에 적힌 타입을 선택하면 됩니다.

> 💡 Title 컬럼은 이미 있으니 새로 만들지 않습니다. 나머지만 추가하면 됩니다.

---

### ① mkqr_Centers (센터 마스터)

| 컬럼명(정확히 입력) | SharePoint 열 유형 | 설정 |
|---|---|---|
| Title | (기본 제공) | 센터명(국문) 저장용, 예: 곤지암센터 |
| Region | 한 줄 텍스트(Single line of text) | |
| CenterNameEn | 한 줄 텍스트 | |
| CorpNameKr | 한 줄 텍스트 | |
| AddressKr | 한 줄 텍스트 | |
| AddressEn | 한 줄 텍스트 | |
| Lat | 숫자(Number) | 소수점 필요 |
| Lng | 숫자(Number) | 소수점 필요 |
| Phone | 한 줄 텍스트 | |
| QrToken | 한 줄 텍스트 | |
| Active | 예/아니요(Yes/No) | 기본값 "예" 권장 |

만든 뒤 확인: 목록 상단 `+ 새 항목`으로 테스트 항목 1개를 넣어보고, 관리자 화면(index.html) → 센터 관리 → "엑셀 일괄 가져오기"로 실제 데이터를 채우세요.

---

### ② mkqr_Staff (선택 — 지금 당장 안 만들어도 시스템은 동작. 추후 "담당자 선택 드롭다운" 기능 추가 시 사용)

| 컬럼명 | 유형 |
|---|---|
| Title | (기본, 담당자 이름) |
| CenterName | 한 줄 텍스트 |
| Phone | 한 줄 텍스트 |
| Email | 한 줄 텍스트 |
| Dept | 한 줄 텍스트 |
| Position | 한 줄 텍스트 |

---

### ③ mkqr_CheckIns (체크인 기록 — Worker가 씀, **가장 중요**)

| 컬럼명(정확히) | SharePoint 열 유형 | 설정 |
|---|---|---|
| Title | (기본) | 자동 생성됨 |
| CenterId | 한 줄 텍스트 | |
| CenterName | 한 줄 텍스트 | |
| VisitorName | 한 줄 텍스트 | |
| VisitorPhone | 한 줄 텍스트 | |
| ServerTimestamp | 날짜 및 시간(Date and time) | "날짜 및 시간 포함" 선택 |
| ClientTimestamp | 한 줄 텍스트 | 참고용 원본(기기 시각), ISO 문자열 그대로 저장 |
| PublicIP | 한 줄 텍스트 | |
| IpLat | 숫자 | |
| IpLng | 숫자 | |
| IpCity | 한 줄 텍스트 | |
| GpsLat | 숫자 | |
| GpsLng | 숫자 | |
| GpsAccuracy | 숫자 | |
| DistanceKm | 숫자 | |
| DistanceSource | 한 줄 텍스트 | |
| VerifyStatus | 선택(Choice) | 선택 항목에 **정상 / 검증필요 / GPS거부 / 중복의심 / 관리자승인** 5개를 한 줄씩 등록 |
| UserAgent | 여러 줄 텍스트(Multiple lines of text) | "일반 텍스트(Plain text)"로 설정 (서식 있는 텍스트로 하면 값에 HTML 껍데기가 씌워짐) |
| DeviceId | 한 줄 텍스트 | localStorage 기반 임의 ID — 브라우저 데이터 삭제 시 초기화됨 |
| DeviceFingerprint | 한 줄 텍스트 | 화면/캔버스/시간대 등을 조합한 해시(64자) — 같은 기기의 반복 체크인 탐지용. localStorage보다 잘 안 지워짐(완전하진 않음) |

**주의**: VerifyStatus를 Choice로 만들 때 "여기에 사용자가 값을 수동으로 입력하도록 허용" 옵션은 꺼두는 걸 권장합니다(오타 방지). 5개 선택지 철자가 코드(`worker.js`, `index.html`)와 **정확히 일치**해야 필터/배지가 정상 동작합니다.

이 목록은 Worker(앱 전용 인증)가 씁니다. 관리자 화면에서 "관리자승인" 버튼을 누르면 그때는 Jimmy님 로그인 세션(MSAL)으로 값이 바뀝니다.

---

### 권한 관련

- `mkqr_Centers`, `mkqr_Staff`는 관리자 화면(MSAL 로그인한 임직원)만 접근하므로 사이트 기본 권한 그대로 두면 됩니다.
- `mkqr_CheckIns`는 **Worker(앱 전용 인증)**가 써야 하므로, 위 "Entra ID 앱 등록" 단계에서 `Grant-PnPAzureADAppSitePermission`으로 이 앱에 STK-DB 사이트 쓰기 권한을 부여하는 걸 잊지 마세요. (사이트 전체 권한이 부담되면 `Sites.Selected`로 이 사이트만 지정하는 걸 권장합니다 — README 상단 절차에 이미 반영되어 있습니다.)

## 설계 메모 (Jimmy 확인 사항)

- **IP 기반 위치는 정확도가 낮습니다** (특히 모바일 통신사 회선은 실제 위치에서 수십km 벗어나는 경우가 흔함). 그래서 판정은 **GPS 우선, IP는 GPS 거부 시 보조 수단**으로 설계했습니다. GPS를 거부하면 자동으로 "GPS거부"로 분류되어 관리자 확인 대상이 됩니다.
- **신원 확인**: 현재는 이름/연락처를 방문자가 직접 입력합니다. 부정 입력 방지가 중요하다면 mkqr_Staff처럼 각 센터 담당자 목록에서 선택하게 하는 방식으로 바꿀 수 있습니다 (용도가 "캠페인 방문 인증"이라 방문자가 매번 다를 수 있어 일단 자유입력으로 설계했습니다).
- **악의적 반복 스캔 방지**: 브라우저 특성(화면/캔버스/시간대 등)을 조합한 `DeviceFingerprint`를 매 체크인마다 수집해서, 같은 센터에 같은 기기가 `DUPLICATE_WINDOW_HOURS`(기본 24시간) 안에 다시 찍으면 이름을 바꿔도 자동으로 "중복의심"으로 분류합니다(`worker.js`의 `isDuplicateDevice`). 다만 이건 어뷰징 난이도를 높이는 수준이지 완벽한 차단은 아닙니다 — 시크릿모드+다른 브라우저 조합 등으로 핑거프린트 자체를 바꿔버리면 못 잡습니다. 더 강하게 막으려면 QR 자체에 만료/회전 토큰을 넣거나 전화번호 인증(SMS OTP)을 추가하는 방법이 있는데, 비용/UX 트레이드오프가 있어 우선 넣지 않았습니다. 필요하시면 추가하겠습니다.
