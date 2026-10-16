/**
 * 帰宅クイズ — 강사 화면 「학생 분석」용 기록 읽기 + 랭킹 + 일련번호 확인 (Google Apps Script)
 *
 * 설치: 기록 시트 → 확장 프로그램 → Apps Script → 이 코드를 통째로 붙여넣기 → 저장
 *       → 배포 → 새 배포 → 유형 「웹 앱」 → 실행: 나 / 액세스 권한: 모든 사용자 → 배포
 *       → 나오는 「웹 앱 URL」과 아래 KEY 를 강사 화면 「학생 분석 → 연결 설정」에 붙여넣기
 *
 * KEY 를 아는 사람만 기록을 볼 수 있습니다. KEY 를 바꾸면 「배포 관리 → 편집 → 새 버전」으로 다시 배포하세요.
 * 코드를 고친 뒤에는 항상 「배포 → 배포 관리 → ✏️ 편집 → 버전: 새 버전 → 배포」 (URL은 그대로 유지됨)
 *
 * ?mode=rank (키 없이 공개): 학생 화면 랭킹용 — 이번 주·이번 달 상위 5명의 이름(석 자, 2026-10-17 사용자 결정)·정답률·문항 수만 돌려줌 (1분마다 새로 계산)
 *   연습: 🎯 스나이퍼(첫 시도 정답률) · 🔥 갓생러(푼 문항 수) / 퇴실 퀴즈: 정답률만. 원외생 기록은 랭킹에서 뺌.
 * ?mode=who&id=6자리 (키 없이 공개, 2026-10-16): 일련번호 확인 — 원생명단 시트에서 찾아 「원생인지 + 가린 이름(김○은)」만 돌려줌 (번호를 마구 넣어 이름을 알아내지 못하게)
 *
 * 원생명단: 선생님 드라이브의 별도 구글 시트 「帰宅クイズ 원생명단」(공유하지 않음). 첫 시트 1행 머리글: 일련번호 · 이름 · 퀴즈반 · … · 상태
 *   한 번이라도 수업을 들은 원생은 지금 반·재원 여부와 상관없이 계속 이용 (2026-10-17 사용자 결정). 막아야 할 때만 상태를 「차단」으로. 행은 지우지 말 것.
 *   기록 시트의 이름 칸에는 원생이면 일련번호(6자리)가 들어옴 → 학생 분석(KEY)에서는 명단의 전체 이름으로 바꿔서 보여 줌.
 */
const KEY = "여기에-비밀-키";            // 실제 키는 저장소에 올리지 말 것
const ROSTER_ID = "여기에-명단-시트-ID";   // 원생명단 시트 주소의 /d/ 와 /edit 사이 값. 저장소에 올리지 말 것

const RANK_EXCLUDE = ["현채원"];          // 랭킹에서 뺄 이름 (선생님 테스트 기록 등)
const RANK_TOP = 5;                       // 몇 위까지 보여줄지
const RANK_MIN = { prac: 50, clsWeek: 1, clsMonth: 3 };   // 랭킹 조건: 연습 50문항 이상 / 퇴실 퀴즈 이번 주 1회·이번 달 3회 이상
const GUEST_CLS = "원외생";               // 원외생 기록의 반 이름 (랭킹 제외)

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.mode === "rank") return out_(rank_(), p.callback);
  if (p.mode === "who") return out_(who_(String(p.id || "")), p.callback);
  if (p.key !== KEY) return out_({ ok: false, error: "key" }, p.callback);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = sheet_();   // 설문 응답 시트(첫 칸이 타임스탬프인 시트), 없으면 첫 시트
  const vals = sh.getDataRange().getValues();
  const tz = ss.getSpreadsheetTimeZone();
  const head = vals.shift().map(String);
  const ni = nameCol_(head), ci = clsCol_(head), R = roster_();
  const rows = vals.filter(r => r[0] !== "").map(r => r.map((v, i) => {
    if (v instanceof Date) {
      // 첫 칸(타임스탬프)은 ISO 시각, 나머지 칸이 날짜로 바뀐 경우(예: "3/10" 점수)는 되돌림
      return i === 0 ? Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ssXXX") : Utilities.formatDate(v, tz, "M/d");
    }
    if (i === ni && String(r[ci] || "").trim() !== GUEST_CLS) { const m = match_(R, v); if (m) return m.name; }   // 일련번호·예전 이름 → 명단 이름
    return v;
  }));
  return out_({ ok: true, head: head, rows: rows, at: new Date().toISOString() }, p.callback);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().find(s => /타임스탬프|Timestamp/i.test(String(s.getRange(1, 1).getValue()))) || ss.getSheets()[0];
}
function nameCol_(head) { const i = head.findIndex((h, k) => k > 0 && /이름|name/i.test(h)); return i >= 0 ? i : 1; }
function clsCol_(head) { const i = head.findIndex((h, k) => k > 0 && /^반|반$|class/i.test(String(h).trim())); return i >= 0 ? i : 2; }

/* ── 원생명단 ── */
function roster_() {
  const cache = CacheService.getScriptCache(), hit = cache.get("roster");
  if (hit) return JSON.parse(hit);
  const R = { ids: {}, base: {} };
  if (!ROSTER_ID || /여기에/.test(ROSTER_ID)) return R;
  const vals = SpreadsheetApp.openById(ROSTER_ID).getSheets()[0].getDataRange().getValues();
  const head = vals.shift().map(h => String(h).trim());
  const col = (re, d) => { const i = head.findIndex(h => re.test(h)); return i >= 0 ? i : d; };
  const C = { id: col(/일련번호|번호/, 0), name: col(/^이름/, 1), st: col(/상태/, 4) };
  vals.forEach(r => {
    const id = String(r[C.id] || "").replace(/\D/g, ""); const name = String(r[C.name] || "").trim();
    if (!/^\d{6}$/.test(id) || !name) return;
    const st = String(r[C.st] || "").trim();
    R.ids[id] = { name: name, out: st === "차단" };
    const b = base_(name); (R.base[b] = R.base[b] || []).push(id);   // 예전 기록(이름만 있는 것) 연결용
  });
  cache.put("roster", JSON.stringify(R), 300);   // 5분 캐시 — 명단을 고치면 5분 안에 반영
  return R;
}
function base_(name) { return String(name).replace(/\(.*?\)/g, "").replace(/\s+/g, ""); }
function mask_(name) {
  const b = base_(name);
  if (b.length <= 1) return b;
  if (b.length === 2) return b[0] + "○";
  return b[0] + "○".repeat(b.length - 2) + b[b.length - 1];
}
function match_(R, v) {   // 일련번호 → 명단 / 예전 이름 → 같은 이름이 명단에 하나뿐이면 그 학생
  const s = String(v || "").trim();
  if (/^\d{6}$/.test(s)) return R.ids[s] ? Object.assign({ id: s }, R.ids[s]) : null;
  const ids = R.base[base_(s)];
  return ids && ids.length === 1 ? Object.assign({ id: ids[0] }, R.ids[ids[0]]) : null;
}

function showName_(R, name) { const b = base_(name); return (R.base[b] || []).length > 1 ? name : b; }

function who_(id) {
  if (!/^\d{6}$/.test(id)) return { ok: false, error: "id" };
  const cache = CacheService.getScriptCache(), mk = "who:" + Math.floor(Date.now() / 60000);
  const n = +(cache.get(mk) || 0) + 1; cache.put(mk, String(n), 120);
  if (n > 120) return { ok: false, error: "busy" };   // 1분에 120번 넘게 물으면 잠시 막음 (번호 마구 넣어 보기 방지)
  const r = roster_().ids[id];
  if (!r || r.out) return { ok: true, member: false };
  return { ok: true, member: true, nm: mask_(r.name) };
}

function rank_() {
  const cache = CacheService.getScriptCache(), CK = "rank3:" + RANK_TOP + ":" + JSON.stringify(RANK_MIN), hit = cache.get(CK);   // 설정이 바뀌면 캐시도 새로
  if (hit) return JSON.parse(hit);
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tz = ss.getSpreadsheetTimeZone();
  const vals = sheet_().getDataRange().getValues(), head = vals.shift().map(h => String(h).trim());
  const col = (re, d) => { const i = head.findIndex((h, k) => k > 0 && re.test(h)); return i >= 0 ? i : d; };
  const C = { name: col(/이름|name/i, 1), cls: col(/^반|반$|class/i, 2), set: col(/세트|범위|set/i, 3), first: col(/첫|정답|점수|first|score/i, 4), time: col(/시간|초\)|time/i, 5) };
  const now = new Date(), day = d => Utilities.formatDate(d, tz, "yyyy-MM-dd");
  const dow = +Utilities.formatDate(now, tz, "u");                       // 1=월 … 7=일
  const weekStart = day(new Date(now.getTime() - (dow - 1) * 864e5)), monthStart = Utilities.formatDate(now, tz, "yyyy-MM-01");
  const skip = new Set(RANK_EXCLUDE.map(n => n.replace(/\s+/g, "")));
  const R = roster_();
  const agg = {};   // [kind|cls|period] -> {key -> {name, f, q, c}}
  const add = (k, key, name, f, q, t) => { const m = agg[k] || (agg[k] = {}); const x = m[key] || (m[key] = { name: name, f: 0, q: 0, c: 0, t: 0 }); x.name = name; x.f += f; x.q += q; x.c++; x.t += t; };
  vals.forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const cls = String(r[C.cls] || "").trim();
    if (cls === GUEST_CLS) return;                                        // 원외생은 랭킹 제외
    const raw = String(r[C.name] || "").trim(), mt = match_(R, raw);
    if (!mt && /^\d{6}$/.test(raw)) return;                               // 명단에 없는 번호
    const key = mt ? mt.id : raw.replace(/\s+/g, "");
    if (!key || skip.has(base_(mt ? mt.name : raw)) || (!mt && !/^[가-힣A-Za-z]{2,}$/.test(key))) return;   // 「.」「기ㅁ현수」처럼 잘못 친 이름은 랭킹에서 제외
    const name = mt ? showName_(R, mt.name) : raw;                        // 랭킹은 이름 석 자 (같은 이름이 명단에 둘 이상이면 (지역)까지)
    let sc = r[C.first]; if (sc instanceof Date) sc = Utilities.formatDate(sc, tz, "M/d");
    const m = String(sc).match(/(\d+)\s*\/\s*(\d+)/); if (!m || !+m[2]) return;
    const prac = cls === "연습" || /^연습 · /.test(String(r[C.set] || ""));
    const d = day(r[0]), f = +m[1], q = +m[2], t = +r[C.time] || 0;
    for (const [per, from] of [["week", weekStart], ["month", monthStart]]) {
      if (d < from) continue;
      if (prac) add("prac|" + per, key, name, f, q, t);
      else { add("cls|전체|" + per, key, name, f, q, t); if (cls) add("cls|" + cls + "|" + per, key, name, f, q, t); }
    }
  });
  const out = { ok: true, at: now.toISOString(), weekStart: weekStart, monthStart: monthStart, min: RANK_MIN, lists: {} };
  Object.keys(agg).forEach(k => {
    const prac = k.indexOf("prac|") === 0;
    const all = Object.keys(agg[k]).map(key => agg[k][key]);
    const need = prac ? RANK_MIN.prac : (/\|week$/.test(k) ? RANK_MIN.clsWeek : RANK_MIN.clsMonth);
    const ok = all.filter(x => (prac ? x.q : x.c) >= need);
    ok.sort((a, b) => b.f / b.q - a.f / a.q || b.q - a.q || a.t / a.q - b.t / b.q);   // 정답률 → 푼 문항 수 → 문항당 시간
    const row = x => ({ name: x.name, acc: Math.round(x.f / x.q * 1000) / 10, q: x.q, c: x.c });
    out.lists[k] = { n: ok.length, total: all.length, top: ok.slice(0, RANK_TOP).map(row) };
    if (prac) {   // 🔥 갓생러: 푼 문항 수 순 (조건 없음), 같으면 정답률 높은 사람이 위
      const qty = all.slice().sort((a, b) => b.q - a.q || b.f / b.q - a.f / a.q);
      out.lists[k + "|qty"] = { n: qty.length, total: all.length, top: qty.slice(0, RANK_TOP).map(row) };
    }
  });
  cache.put(CK, JSON.stringify(out), 60);    // 1분 캐시
  return out;
}

function out_(obj, cb) {
  const s = JSON.stringify(obj);
  if (cb && /^[\w.]+$/.test(cb)) return ContentService.createTextOutput(cb + "(" + s + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
