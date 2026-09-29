/**
 * 帰宅クイズ — 강사 화면 「학생 분석」용 기록 읽기 (Google Apps Script)
 *
 * 설치: 기록 시트 → 확장 프로그램 → Apps Script → 이 코드를 통째로 붙여넣기 → 저장
 *       → 배포 → 새 배포 → 유형 「웹 앱」 → 실행: 나 / 액세스 권한: 모든 사용자 → 배포
 *       → 나오는 「웹 앱 URL」과 아래 KEY 를 강사 화면 「학생 분석 → 연결 설정」에 붙여넣기
 *
 * KEY 를 아는 사람만 기록을 볼 수 있습니다. KEY 를 바꾸면 「배포 관리 → 편집 → 새 버전」으로 다시 배포하세요.
 * 코드를 고친 뒤에는 항상 「배포 → 배포 관리 → ✏️ 편집 → 버전: 새 버전 → 배포」 (URL은 그대로 유지됨)
 *
 * ?mode=rank (키 없이 공개): 학생 화면 랭킹용 — 이번 주·이번 달 상위 5명의 이름·정답률·문항 수만 돌려줌 (1분마다 새로 계산)
 *   연습: 🎯 스나이퍼(첫 시도 정답률) · 🔥 갓생러(푼 문항 수) / 퇴실 퀴즈: 정답률만
 */
const KEY = "여기에-비밀-키";   // 실제 키는 저장소에 올리지 말 것

const RANK_EXCLUDE = ["현채원"];          // 랭킹에서 뺄 이름 (선생님 테스트 기록 등)
const RANK_TOP = 5;                       // 몇 위까지 보여줄지
const RANK_MIN = { prac: 50, clsWeek: 1, clsMonth: 3 };   // 랭킹 조건: 연습 50문항 이상 / 퇴실 퀴즈 이번 주 1회·이번 달 3회 이상

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.mode === "rank") return out_(rank_(), p.callback);
  if (p.key !== KEY) return out_({ ok: false, error: "key" }, p.callback);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = sheet_();   // 설문 응답 시트(첫 칸이 타임스탬프인 시트), 없으면 첫 시트
  const vals = sh.getDataRange().getValues();
  const tz = ss.getSpreadsheetTimeZone();
  const head = vals.shift().map(String);
  const rows = vals.filter(r => r[0] !== "").map(r => r.map((v, i) => {
    if (v instanceof Date) {
      // 첫 칸(타임스탬프)은 ISO 시각, 나머지 칸이 날짜로 바뀐 경우(예: "3/10" 점수)는 되돌림
      return i === 0 ? Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ssXXX") : Utilities.formatDate(v, tz, "M/d");
    }
    return v;
  }));
  return out_({ ok: true, head: head, rows: rows, at: new Date().toISOString() }, p.callback);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().find(s => /타임스탬프|Timestamp/i.test(String(s.getRange(1, 1).getValue()))) || ss.getSheets()[0];
}

function rank_() {
  const cache = CacheService.getScriptCache(), hit = cache.get("rank");
  if (hit) return JSON.parse(hit);
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tz = ss.getSpreadsheetTimeZone();
  const vals = sheet_().getDataRange().getValues(), head = vals.shift().map(h => String(h).trim());
  const col = (re, d) => { const i = head.findIndex((h, k) => k > 0 && re.test(h)); return i >= 0 ? i : d; };
  const C = { name: col(/이름|name/i, 1), cls: col(/^반|반$|class/i, 2), set: col(/세트|범위|set/i, 3), first: col(/첫|정답|점수|first|score/i, 4), time: col(/시간|초\)|time/i, 5) };
  const now = new Date(), day = d => Utilities.formatDate(d, tz, "yyyy-MM-dd");
  const dow = +Utilities.formatDate(now, tz, "u");                       // 1=월 … 7=일
  const weekStart = day(new Date(now.getTime() - (dow - 1) * 864e5)), monthStart = Utilities.formatDate(now, tz, "yyyy-MM-01");
  const skip = new Set(RANK_EXCLUDE.map(n => n.replace(/\s+/g, "")));
  const agg = {};   // [kind|cls|period] -> {key -> {name, f, q, c}}
  const add = (k, key, name, f, q, t) => { const m = agg[k] || (agg[k] = {}); const x = m[key] || (m[key] = { name: name, f: 0, q: 0, c: 0, t: 0 }); x.name = name; x.f += f; x.q += q; x.c++; x.t += t; };
  vals.forEach(r => {
    if (!(r[0] instanceof Date)) return;
    const name = String(r[C.name] || "").trim(), key = name.replace(/\s+/g, "");
    if (!key || skip.has(key) || !/^[가-힣A-Za-z]{2,}$/.test(key)) return;   // 「.」「기ㅁ현수」처럼 잘못 친 이름은 랭킹에서 제외
    let sc = r[C.first]; if (sc instanceof Date) sc = Utilities.formatDate(sc, tz, "M/d");
    const m = String(sc).match(/(\d+)\s*\/\s*(\d+)/); if (!m || !+m[2]) return;
    const cls = String(r[C.cls] || "").trim(), prac = cls === "연습" || /^연습 · /.test(String(r[C.set] || ""));
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
  cache.put("rank", JSON.stringify(out), 60);    // 1분 캐시
  return out;
}

function out_(obj, cb) {
  const s = JSON.stringify(obj);
  if (cb && /^[\w.]+$/.test(cb)) return ContentService.createTextOutput(cb + "(" + s + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
