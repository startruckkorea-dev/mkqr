/**
 * MKQR Worker
 * ------------------------------------------------------------------
 * checkin.html(공개, 로그인 없음)이 호출하는 API. 로그인한 사람이 없으므로
 * MSAL 사용자 토큰이 없고, 대신 Azure AD "앱 자격증명(Client Credentials)"으로
 * Graph API에 접근한다. client secret은 반드시 이 Worker(서버)에만 보관하고
 * 절대 GitHub Pages(정적 HTML)에 넣지 않는다.
 *
 * 필요한 사전 준비 (Jimmy 작업):
 * 1) Entra ID(Azure AD) 새 App 등록 (기존 APS-MasterData-App과 별도로 만드는 걸 권장 —
 *    이 앱은 "앱 자격증명"으로 SharePoint에 쓰기 때문에 사용자 위임 앱과 분리하는 게 안전함)
 * 2) API 권한 → Application permissions → Sites.Selected 추가 → 관리자 동의
 *    (Sites.ReadWrite.All 전체보다 Sites.Selected + 특정 사이트만 권한부여가 더 안전)
 * 3) PnP PowerShell 또는 Graph API로 STK-DB 사이트에 이 App의 Sites.Selected 쓰기 권한 부여
 *    (Grant-PnPAzureADAppSitePermission 참고 — 기존 STK 인트라 백업 자동화 때 쓰신 PnP 모듈 재사용 가능)
 * 4) 인증서/시크릿 → 새 클라이언트 시크릿 생성 (만료일 등록 필수, 캘린더에 갱신 알림 권장)
 * 5) wrangler secret put 으로 아래 값들을 이 Worker에 등록:
 *      wrangler secret put MK_CLIENT_ID
 *      wrangler secret put MK_CLIENT_SECRET
 *      wrangler secret put MK_TENANT_ID
 *      wrangler secret put MK_RECAPTCHA_SECRET   (reCAPTCHA v3 비밀 키, 선택)
 * 6) SharePoint STK-DB 사이트에 리스트 3개 생성 (컬럼 스키마는 대화 중 안내받은 내용 참고):
 *      mkqr_Centers (RedirectUrl 컬럼 포함) / mkqr_Staff / mkqr_CheckIns (RecaptchaScore 컬럼 포함, 숫자형)
 * 7) wrangler.toml의 route/도메인 설정 후 `wrangler deploy`
 * 8) index.html의 CHECKIN_API_BASE, checkin.html의 API_BASE를 실제 배포 URL로 교체
 */

const GRAPH = "https://graph.microsoft.com/v1.0";
const SITE_PATH = "startruckkorea.sharepoint.com:/sites/STK-DB:";
const VERIFY_RADIUS_M_DEFAULT = 10000; // 센터에 별도 설정이 없을 때 쓰는 기본 허용 반경(m)
const EXTRA_REVIEW_DISTANCE_M = 300;    // 센터 반경 설정과 별개로 항상 적용되는 촘촘한 기준
const EXTRA_REVIEW_RECAPTCHA_MIN = 0.5; // 이 값 이상(=사람처럼 행동)인데도 300m 넘게 떨어졌으면 관리자 확인 필요
const REPEAT_VISIT_REVIEW_COUNT = 5;    // 같은 센터를 같은 기기가 이 횟수 이상(누적, 기간제한 없음) 찍으면 무조건 관리자 확인 필요
                                         // - 리워드(방문 실적 포상) 대상 센터를 특정 직원이 반복 태깅해서 부풀리는 것을 막기 위한 규칙
const RATE_LIMIT_WINDOW_MS = 60 * 1000;        // 1분
const RATE_LIMIT_COUNT = 5;                     // 이 시간 안에 이 횟수 초과 접근하면 차단
const RATE_LIMIT_BLOCK_MS = 30 * 60 * 1000;     // 30분 차단

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*", // 필요시 checkin.html이 배포된 정확한 origin으로 좁히세요
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json", ...CORS_HEADERS },
  });
}

// ---- 앱 전용(client credentials) Graph 토큰, 만료 전까지 메모리 캐시 ----
let _cachedToken = null;
let _cachedTokenExp = 0;
async function getAppToken(env) {
  const now = Date.now() / 1000;
  if (_cachedToken && now < _cachedTokenExp - 60) return _cachedToken;
  const url = `https://login.microsoftonline.com/${env.MK_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: env.MK_CLIENT_ID,
    client_secret: env.MK_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(url, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body });
  const data = await res.json();
  if (!res.ok) throw new Error("앱 토큰 발급 실패: " + JSON.stringify(data));
  _cachedToken = data.access_token;
  _cachedTokenExp = now + data.expires_in;
  return _cachedToken;
}

let _siteIdCache = null;
async function getSiteId(env) {
  if (_siteIdCache) return _siteIdCache;
  const token = await getAppToken(env);
  const res = await fetch(`${GRAPH}/sites/${SITE_PATH}`, { headers: { Authorization: "Bearer " + token } });
  const data = await res.json();
  if (!res.ok) throw new Error("사이트 조회 실패: " + JSON.stringify(data));
  _siteIdCache = data.id;
  return _siteIdCache;
}

const _listIdCache = {};
async function getListId(env, listName) {
  if (_listIdCache[listName]) return _listIdCache[listName];
  const token = await getAppToken(env);
  const siteId = await getSiteId(env);
  const res = await fetch(`${GRAPH}/sites/${siteId}/lists?$filter=displayName eq '${listName}'`, { headers: { Authorization: "Bearer " + token } });
  const data = await res.json();
  if (!res.ok || !data.value?.length) throw new Error(`목록을 찾을 수 없음: ${listName}`);
  _listIdCache[listName] = data.value[0].id;
  return _listIdCache[listName];
}

async function listItems(env, listName, filterQuery = "") {
  const token = await getAppToken(env);
  const siteId = await getSiteId(env);
  const listId = await getListId(env, listName);
  let items = [];
  let url = `${GRAPH}/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200${filterQuery}`;
  while (url) {
    const res = await fetch(url, {
      headers: { Authorization: "Bearer " + token, Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" },
    });
    const data = await res.json();
    if (!res.ok) throw new Error(JSON.stringify(data));
    items = items.concat(data.value || []);
    url = data["@odata.nextLink"] || null;
  }
  return items;
}

// 체크인 기록이 수천 건 쌓여도 매번 전체를 긁지 않도록, ServerTimestamp 기준으로 최근 것만 서버측 필터링해서 가져온다.
// mkqr_CheckIns 목록의 ServerTimestamp 컬럼에 "인덱스"를 걸어두는 걸 권장 (목록 설정 → 인덱스 걸린 열 → 만들기).
// 인덱스가 없어도 Prefer 헤더 덕분에 동작은 하지만, 목록이 5,000건을 넘어가면 느려지거나 실패할 수 있다.
async function listRecentCheckins(env, sinceMs) {
  const sinceIso = new Date(sinceMs).toISOString();
  return listItems(env, "mkqr_CheckIns", `&$filter=fields/ServerTimestamp ge '${sinceIso}'`);
}

// 같은 센터를 같은 기기가 지금까지(기간 제한 없이) 몇 번 찍었는지 조회 - 리워드 부정사용 방지용.
// CenterId/DeviceFingerprint는 인덱스가 없는 컬럼이라 HonorNonIndexedQueriesWarningMayFailRandomly로 조회하며,
// 조회 자체가 실패해도(네트워크 오류 등) 체크인 자체를 막지 않기 위해 0으로 처리한다(fail-open).
async function countDeviceCenterVisits(env, centerId, fingerprint) {
  if (!fingerprint || !centerId) return 0;
  try {
    const items = await listItems(env, "mkqr_CheckIns",
      `&$filter=fields/CenterId eq '${centerId}' and fields/DeviceFingerprint eq '${fingerprint}'`);
    return items.filter((it) => it.fields.VerifyStatus !== "차단해제").length;
  } catch (e) {
    return 0;
  }
}

async function createItem(env, listName, fields) {
  const token = await getAppToken(env);
  const siteId = await getSiteId(env);
  const listId = await getListId(env, listName);
  const res = await fetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items`, {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ fields }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(JSON.stringify(data));
  return data;
}

function distanceKm(lat1, lng1, lat2, lng2) {
  const R = 6371;
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

async function ipGeolocate(ip) {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 5000);
    const res = await fetch(`https://ipapi.co/${ip}/json/`, { signal: controller.signal });
    clearTimeout(timer);
    if (!res.ok) return null;
    const d = await res.json();
    if (d.error || d.latitude == null) return null;
    // 서울/인천/부산/대구 등 광역시는 도시명=지역명이 같아서 그냥 합치면 "Seoul Seoul"처럼 중복 표시됨 - 같으면 한 번만.
    const cityRegionParts = [d.city, d.region].filter(Boolean);
    const uniqueParts = cityRegionParts.filter((v, i) => cityRegionParts.indexOf(v) === i);
    return { lat: d.latitude, lng: d.longitude, city: uniqueParts.join(" ") };
  } catch (e) {
    return null;
  }
}

// reCAPTCHA v3 서버측 검증. 토큰이 없거나 시크릿이 아직 설정 안 됐거나 Google 쪽에 문제가 있어도
// 체크인 자체를 막지는 않는다 (광고차단기 등으로 스크립트가 막히는 경우가 흔해서, 실패해도 그냥
// score를 null로 두고 넘어가고 다른 신호(GPS/IP/기기 반복)로 계속 검증한다).
async function verifyRecaptcha(env, token, remoteIp) {
  if (!token || !env.MK_RECAPTCHA_SECRET) return null;
  try {
    const body = new URLSearchParams({ secret: env.MK_RECAPTCHA_SECRET, response: token, remoteip: remoteIp || "" });
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 5000);
    const res = await fetch("https://www.google.com/recaptcha/api/siteverify", {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body, signal: controller.signal,
    });
    clearTimeout(timer);
    const data = await res.json();
    if (!data.success) return null;
    return typeof data.score === "number" ? data.score : null;
  } catch (e) {
    return null;
  }
}

async function findCenterByToken(env, token) {
  // SharePoint 서식 필드는 대소문자/공백에 예민할 수 있어 전체 조회 후 JS로 비교 (센터 수가 적어 부담 없음)
  const items = await listItems(env, "mkqr_Centers");
  return items.find((it) => it.fields.QrToken === token);
}

// 1분 안에 같은 기기가 같은 센터에 5회 초과 접근했는지 확인하고, 그렇다면 30분 차단 상태를 계산한다.
// 차단 여부는 저장된 상태값이 아니라 mkqr_CheckIns 기록만으로 매번 다시 계산한다:
//  1) 관리자가 "차단해제" 마커를 남기면, 그 시점 이후의 기록만 갖고 판단한다.
//  2) 남은 기록을 시간순으로 훑으면서, 어느 시점이든 최근 1분 안에 5건이 몰려있으면
//     그 5번째 기록 시각 + 30분을 차단 해제 시각(blockUntil)으로 잡는다 (가장 늦은 값 사용).
// 전체 이력이 아니라 최근 2시간치만 조회한다 (차단 판정에 필요한 범위는 최대 30분+여유분이면 충분,
// 목록이 아무리 커져도 이 조회량은 늘어나지 않는다).
async function computeBlockStatus(env, centerId, fingerprint) {
  if (!fingerprint) return { blocked: false };
  const items = await listRecentCheckins(env, Date.now() - 2 * 3600 * 1000);
  const sameDevice = items.filter((it) => it.fields.CenterId === String(centerId) && it.fields.DeviceFingerprint === fingerprint);

  const lastOverride = sameDevice
    .filter((it) => it.fields.VerifyStatus === "차단해제")
    .reduce((max, it) => Math.max(max, new Date(it.fields.ServerTimestamp || 0).getTime()), 0);

  const events = sameDevice
    .filter((it) => it.fields.VerifyStatus !== "차단해제")
    .map((it) => new Date(it.fields.ServerTimestamp || 0).getTime())
    .filter((t) => t > lastOverride)
    .sort((a, b) => a - b);

  let blockUntil = 0;
  for (let i = 0; i < events.length; i++) {
    let count = 0;
    for (let j = i; j >= 0 && events[i] - events[j] <= RATE_LIMIT_WINDOW_MS; j--) count++;
    if (count >= RATE_LIMIT_COUNT) blockUntil = Math.max(blockUntil, events[i] + RATE_LIMIT_BLOCK_MS);
  }

  const now = Date.now();
  if (blockUntil > now) return { blocked: true, remainingMinutes: Math.ceil((blockUntil - now) / 60000) };
  return { blocked: false };
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") return new Response(null, { headers: CORS_HEADERS });

    try {
      if (url.pathname === "/api/center" && request.method === "GET") {
        const token = url.searchParams.get("t");
        if (!token) return json({ error: "잘못된 요청입니다" }, 400);
        const center = await findCenterByToken(env, token);
        if (!center || center.fields.Active === false) return json({ error: "유효하지 않은 QR 코드입니다" }, 404);
        return json({ name: center.fields.Title, redirectUrl: center.fields.RedirectUrl || "" });
      }

      if (url.pathname === "/api/checkin" && request.method === "POST") {
        const body = await request.json();
        const { token, gps, clientTimestamp, deviceId, deviceFingerprint, userAgent, recaptchaToken } = body;
        if (!token) return json({ error: "잘못된 요청입니다" }, 400);

        const center = await findCenterByToken(env, token);
        if (!center || center.fields.Active === false) return json({ error: "유효하지 않은 QR 코드입니다" }, 404);
        const c = center.fields;

        // 1분 안에 5회 초과 접근이면 기록을 남기지 않고 즉시 차단 응답 (스팸 로그 방지)
        const block = await computeBlockStatus(env, center.id, deviceFingerprint);
        if (block.blocked) {
          return json({
            blocked: true,
            error: `동일 기기 및 동일 위치 등으로 다수 접근이 확인되었습니다. ${block.remainingMinutes}분 뒤 또는 관리자의 해당 기기 설정 허용 이후 다시 QR을 통한 링크 접속이 가능합니다.`,
          }, 429);
        }

        // Cloudflare가 넘겨주는 실제 접속 IP — 클라이언트가 조작할 수 없는 신뢰 가능한 공인 IP
        const publicIp = request.headers.get("cf-connecting-ip") || "unknown";
        const cfCityRegionParts = request.cf ? [request.cf.city, request.cf.region].filter(Boolean) : [];
        const cfCity = [...new Set(cfCityRegionParts)].join(" ");

        // 서로 관련 없는 외부 호출(IP 위치조회, reCAPTCHA 검증)은 순서대로 기다리지 않고 동시에 실행해서 지연을 줄인다.
        const [ipLoc, recaptchaScore] = await Promise.all([
          publicIp !== "unknown" ? ipGeolocate(publicIp) : Promise.resolve(null),
          verifyRecaptcha(env, recaptchaToken, publicIp),
        ]);
        const ipCity = ipLoc?.city || cfCity || "";

        let distKm = null, source = "없음", verifyStatus = "검증필요", gpsUsed = false;
        const hasCenterCoord = c.Lat != null && c.Lng != null;
        const radiusKm = (c.VerifyRadiusM != null ? Number(c.VerifyRadiusM) : VERIFY_RADIUS_M_DEFAULT) / 1000;

        if (gps && gps.lat != null && hasCenterCoord) {
          distKm = distanceKm(gps.lat, gps.lng, c.Lat, c.Lng);
          source = "GPS";
          gpsUsed = true;
          verifyStatus = distKm <= radiusKm ? "정상" : "검증필요";
        } else if (ipLoc && hasCenterCoord) {
          distKm = distanceKm(ipLoc.lat, ipLoc.lng, c.Lat, c.Lng);
          source = "IP";
          verifyStatus = distKm <= radiusKm ? "정상" : "검증필요";
        } else if (!gps) {
          verifyStatus = "GPS거부";
        }

        // reCAPTCHA 점수가 낮으면(봇 의심) 위치가 맞았더라도 검증필요로 내림
        if (recaptchaScore != null && recaptchaScore < 0.3 && verifyStatus === "정상") {
          verifyStatus = "검증필요";
        }

        // 센터 허용 반경과는 별개로 항상 적용되는 규칙: 300m 넘게 떨어졌는데 reCAPTCHA 점수가
        // 사람처럼 행동했다고 볼 만큼 높으면(0.5 이상) — 반경 안이라 "정상"으로 판정됐더라도
        // 한 번 더 관리자가 눈으로 확인하도록 "검증필요"로 내린다.
        if (verifyStatus === "정상" && distKm != null && distKm * 1000 >= EXTRA_REVIEW_DISTANCE_M
            && recaptchaScore != null && recaptchaScore >= EXTRA_REVIEW_RECAPTCHA_MIN) {
          verifyStatus = "검증필요";
        }

        // 리워드 부정사용 방지: 이번 건까지 포함해 같은 센터+기기 누적 방문이 REPEAT_VISIT_REVIEW_COUNT회
        // 이상이면 위치·reCAPTCHA가 전부 정상이어도 무조건 관리자 확인이 필요하도록 내린다.
        // (1분 5회 초과 즉시차단 규칙과는 별개 - 그건 짧은 시간 내 반복 접근을, 이건 기간 제한 없는 누적 반복 방문을 잡는다)
        if (verifyStatus === "정상") {
          const priorVisits = await countDeviceCenterVisits(env, String(center.id), deviceFingerprint);
          if (priorVisits + 1 >= REPEAT_VISIT_REVIEW_COUNT) {
            verifyStatus = "검증필요";
          }
        }

        const now = new Date().toISOString();
        const fields = {
          Title: `${c.Title} ${now.slice(0, 16).replace("T", " ")}`,
          CenterId: String(center.id),
          CenterName: c.Title,
          ServerTimestamp: now,
          ClientTimestamp: clientTimestamp || "",
          PublicIP: publicIp,
          IpLat: ipLoc?.lat ?? null,
          IpLng: ipLoc?.lng ?? null,
          IpCity: ipCity,
          GpsLat: gps?.lat ?? null,
          GpsLng: gps?.lng ?? null,
          GpsAccuracy: gps?.accuracy ?? null,
          DistanceKm: distKm != null ? Math.round(distKm * 100) / 100 : null,
          DistanceSource: source,
          VerifyStatus: verifyStatus,
          UserAgent: userAgent || "",
          DeviceId: deviceId || "",
          DeviceFingerprint: deviceFingerprint || "",
          RecaptchaScore: recaptchaScore,
        };
        await createItem(env, "mkqr_CheckIns", fields);

        return json({ verifyStatus, distanceKm: fields.DistanceKm, gpsUsed, redirectUrl: c.RedirectUrl || "" });
      }

      return json({ error: "not found" }, 404);
    } catch (e) {
      return json({ error: e.message }, 500);
    }
  },
};
