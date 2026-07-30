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

## SharePoint 리스트 스키마

### mkqr_Centers (센터 마스터)
| 컬럼 | 타입 | 비고 |
|---|---|---|
| Title | 한 줄 텍스트 | 센터명(국문), 예: 곤지암센터 |
| Region | 한 줄 텍스트 | |
| CenterNameEn | 한 줄 텍스트 | |
| CorpNameKr | 한 줄 텍스트 | |
| AddressKr | 한 줄 텍스트 | |
| AddressEn | 한 줄 텍스트 | |
| Lat / Lng | 숫자 | 좌표 (관리자 화면 "좌표 찾기"로 채움) |
| Phone | 한 줄 텍스트 | |
| QrToken | 한 줄 텍스트 | 센터별 고유 랜덤 토큰, QR URL에 사용 |
| Active | 예/아니요 | QR 활성화 여부 |

### mkqr_Staff (선택 — 센터별 담당자, 추후 신원확인 드롭다운용)
| 컬럼 | 타입 |
|---|---|
| Title | 유저명 |
| CenterName | 한 줄 텍스트 |
| Phone / Email / Dept / Position | 한 줄 텍스트 |

### mkqr_CheckIns (체크인 기록, Worker가 씀)
| 컬럼 | 타입 | 비고 |
|---|---|---|
| Title | 한 줄 텍스트 | 자동 생성 |
| CenterId / CenterName | 한 줄 텍스트 | |
| VisitorName / VisitorPhone | 한 줄 텍스트 | |
| ServerTimestamp | 날짜/시간 | Worker 수신 시각(신뢰 가능) |
| ClientTimestamp | 한 줄 텍스트 | 참고용(기기 시각, 조작 가능) |
| PublicIP | 한 줄 텍스트 | Cloudflare가 확인한 실제 접속 IP |
| IpLat / IpLng / IpCity | 숫자/텍스트 | IP 기반 위치(보조 신호, 정확도 낮음) |
| GpsLat / GpsLng / GpsAccuracy | 숫자 | 브라우저 GPS(1차 판정 기준), 동의 거부 시 비어있음 |
| DistanceKm | 숫자 | 센터 좌표 기준 산출 거리 |
| DistanceSource | 한 줄 텍스트 | GPS / IP / 없음 |
| VerifyStatus | 선택(Choice) | 정상 / 검증필요 / GPS거부 / 관리자승인 |
| UserAgent | 여러 줄 텍스트 | 기기/브라우저 |
| DeviceId | 한 줄 텍스트 | 브라우저 localStorage 기반 임의 ID (재방문 기기 식별용) |

## 설계 메모 (Jimmy 확인 사항)

- **IP 기반 위치는 정확도가 낮습니다** (특히 모바일 통신사 회선은 실제 위치에서 수십km 벗어나는 경우가 흔함). 그래서 판정은 **GPS 우선, IP는 GPS 거부 시 보조 수단**으로 설계했습니다. GPS를 거부하면 자동으로 "GPS거부"로 분류되어 관리자 확인 대상이 됩니다.
- **신원 확인**: 현재는 이름/연락처를 방문자가 직접 입력합니다. 부정 입력 방지가 중요하다면 mkqr_Staff처럼 각 센터 담당자 목록에서 선택하게 하는 방식으로 바꿀 수 있습니다 (용도가 "캠페인 방문 인증"이라 방문자가 매번 다를 수 있어 일단 자유입력으로 설계했습니다).
- **QR 부정 재사용(스크린샷을 다른 곳에서 제출)**까지 막으려면 QR 자체에 만료/회전 토큰을 넣는 방법이 있는데, 우선순위가 아니라면 이번 버전엔 넣지 않았습니다. 필요하시면 추가하겠습니다.
