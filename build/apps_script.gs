/**
 * 帰宅クイズ — 강사 화면 「학생 분석」용 기록 읽기 (Google Apps Script)
 *
 * 설치: 기록 시트 → 확장 프로그램 → Apps Script → 이 코드를 통째로 붙여넣기 → 저장
 *       → 배포 → 새 배포 → 유형 「웹 앱」 → 실행: 나 / 액세스 권한: 모든 사용자 → 배포
 *       → 나오는 「웹 앱 URL」과 아래 KEY 를 강사 화면 「학생 분석 → 연결 설정」에 붙여넣기
 *
 * KEY 를 아는 사람만 기록을 볼 수 있습니다. KEY 를 바꾸면 「배포 관리 → 편집 → 새 버전」으로 다시 배포하세요.
 */
const KEY = "여기에-비밀-키";   // 실제 키는 저장소에 올리지 말 것

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== KEY) return out_({ ok: false, error: "key" }, p.callback);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 설문 응답 시트(첫 칸이 타임스탬프인 시트)를 찾고, 없으면 첫 시트
  const sh = ss.getSheets().find(s => /타임스탬프|Timestamp/i.test(String(s.getRange(1, 1).getValue()))) || ss.getSheets()[0];
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

function out_(obj, cb) {
  const s = JSON.stringify(obj);
  if (cb && /^[\w.]+$/.test(cb)) return ContentService.createTextOutput(cb + "(" + s + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
