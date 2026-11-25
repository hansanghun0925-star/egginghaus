/**
 * Egginghaus 학생 기록 → 구글 스프레드시트
 * 붙여 넣을 곳: 스프레드시트 > 확장 프로그램 > Apps Script
 * 배포: 배포 > 새 배포 > 유형 "웹 앱" > 실행 계정 "나" > 액세스 권한 "모든 사용자"
 */
const SHEET = '기록';
const COLS = ['ts', 'sid', 'sess', 'eid', 'q', 'stage', 'skill_id', 'skill', 'result', 'level', 'conf', 'answer', 'sec', 'conds', 'pats', 'extra'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 같은 기록이 두 번 오면(재전송) 한 번만 적기
function recentEids_(sh) {
  const n = sh.getLastRow() - 1;
  if (n <= 0) return new Set();
  const take = Math.min(n, 5000);
  return new Set(sh.getRange(sh.getLastRow() - take + 1, 4, take, 1).getValues().map(r => String(r[0])));
}

// 수식으로 해석되지 않게
function cell_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') v = JSON.stringify(v);
  if (typeof v === 'string' && /^[=+@]/.test(v)) return "'" + v;
  return v;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents);
    const list = Array.isArray(data) ? data : [data];
    const sh = sheet_();
    const seen = recentEids_(sh);
    const rows = list
      .filter(r => r && r.eid && !seen.has(String(r.eid)))
      .map(r => COLS.map(k => cell_(r[k])));
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, COLS.length).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, added: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 주소가 살아 있는지 브라우저로 열어 확인할 때
function doGet() {
  return ContentService.createTextOutput('Egginghaus 기록 받는 중');
}
