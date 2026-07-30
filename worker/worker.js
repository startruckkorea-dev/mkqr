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
 * 5) wrangler secret put 으로 아래 3개 값을 이 Worker에 등록:
 *      wrangler secret put MK_CLIENT_ID
 *      wrangler secret put MK_CLIENT_SECRET
 *      wrangler secret put MK_TENANT_ID
 * 6) SharePoint STK-DB 사이트에 리스트 3개 생성 (컬럼명은 README.md 스키마 참고):
 *      mkqr_Centers / mkqr_Staff / mkqr_CheckIns
 * 7) wrangler.toml의 route/도메인 설정 후 `wrangler deploy`
 * 8) index.html의 CHECKIN_API_BASE, checkin.html의 API_BASE를 실제 배포 URL로 교체
 */

const GRAPH = "https://graph.microsoft.com/v1.0";
const SITE_PATH = "startruckkorea.sharepoint.com:/sites/STK-DB:";
const VERIFY_RADIUS_KM = 10;
const DUPLICATE_WINDOW_HOURS = 24; // 같은 기기 핑거프린트가 같은 센터에 이 시간 안에 다시 찍으면 "중복의심"

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
    const res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    const data = await res.json();
    if (!res.ok) throw new Error(JSON.stringify(data));
    items = items.concat(data.value || []);
    url = data["@odata.nextLink"] || null;
  }
  return items;
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
    const res = await fetch(`https://ipapi.co/${ip}/json/`);
    if (!res.ok) return null;
    const d = await res.json();
    if (d.error || d.latitude == null) return null;
    return { lat: d.latitude, lng: d.longitude, city: [d.city, d.region].filter(Boolean).join(" ") };
  } catch (e) {
    return null;
  }
}

async function findCenterByToken(env, token) {
  // SharePoint 서식 필드는 대소문자/공백에 예민할 수 있어 전체 조회 후 JS로 비교 (센터 수가 적어 부담 없음)
  const items = await listItems(env, "mkqr_Centers");
  return items.find((it) => it.fields.QrToken === token);
}

// 최근 체크인 기록을 훑어서 같은 센터에 같은 기기 핑거프린트가 이미 다녀갔는지 확인.
// SharePoint List API로 날짜/텍스트 복합 필터를 걸기 번거로워, 전체를 가져온 뒤 JS에서 걸러낸다
// (체크인 건수가 아주 많아지면 $filter=CenterId eq '...' 정도로 서버측 필터를 추가하는 게 좋음).
async function isDuplicateDevice(env, centerId, fingerprint) {
  if (!fingerprint) return false;
  const items = await listItems(env, "mkqr_CheckIns");
  const cutoff = Date.now() - DUPLICATE_WINDOW_HOURS * 3600 * 1000;
  return items.some((it) => {
    const f = it.fields;
    if (f.CenterId !== String(centerId)) return false;
    if (f.DeviceFingerprint !== fingerprint) return false;
    const t = f.ServerTimestamp ? new Date(f.ServerTimestamp).getTime() : 0;
    return t >= cutoff;
  });
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
        return json({ name: center.fields.Title });
      }

      if (url.pathname === "/api/checkin" && request.method === "POST") {
        const body = await request.json();
        const { token, name, phone, gps, clientTimestamp, deviceId, deviceFingerprint, userAgent } = body;
        if (!token || !name) return json({ error: "이름을 입력해주세요" }, 400);

        const center = await findCenterByToken(env, token);
        if (!center || center.fields.Active === false) return json({ error: "유효하지 않은 QR 코드입니다" }, 404);
        const c = center.fields;

        // Cloudflare가 넘겨주는 실제 접속 IP — 클라이언트가 조작할 수 없는 신뢰 가능한 공인 IP
        const publicIp = request.headers.get("cf-connecting-ip") || "unknown";
        const cfCity = request.cf ? [request.cf.city, request.cf.region].filter(Boolean).join(" ") : "";

        let ipLoc = null;
        if (publicIp !== "unknown") ipLoc = await ipGeolocate(publicIp);
        const ipCity = ipLoc?.city || cfCity || "";

        let distKm = null, source = "없음", verifyStatus = "검증필요", gpsUsed = false;
        const hasCenterCoord = c.Lat != null && c.Lng != null;

        if (gps && gps.lat != null && hasCenterCoord) {
          distKm = distanceKm(gps.lat, gps.lng, c.Lat, c.Lng);
          source = "GPS";
          gpsUsed = true;
          verifyStatus = distKm <= VERIFY_RADIUS_KM ? "정상" : "검증필요";
        } else if (ipLoc && hasCenterCoord) {
          distKm = distanceKm(ipLoc.lat, ipLoc.lng, c.Lat, c.Lng);
          source = "IP";
          verifyStatus = distKm <= VERIFY_RADIUS_KM ? "정상" : "검증필요";
        } else if (!gps) {
          verifyStatus = "GPS거부";
        }

        // 위치 판정과는 별개로, 같은 기기가 최근에 같은 센터를 이미 찍었으면 최우선으로 "중복의심" 처리
        const dup = await isDuplicateDevice(env, center.id, deviceFingerprint);
        if (dup) verifyStatus = "중복의심";

        const now = new Date().toISOString();
        const fields = {
          Title: `${c.Title} ${now.slice(0, 16).replace("T", " ")}`,
          CenterId: String(center.id),
          CenterName: c.Title,
          VisitorName: name,
          VisitorPhone: phone || "",
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
        };
        await createItem(env, "mkqr_CheckIns", fields);

        return json({ verifyStatus, distanceKm: fields.DistanceKm, gpsUsed });
      }

      return json({ error: "not found" }, 404);
    } catch (e) {
      return json({ error: e.message }, 500);
    }
  },
};
