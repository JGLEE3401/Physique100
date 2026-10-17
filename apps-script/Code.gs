// Physique100 ↔ 구글시트 동기화용 Apps Script (시트에 바인딩된 스크립트로 붙여넣기)
// 토큰은 코드에 쓰지 말고 프로젝트 설정 → 스크립트 속성에 TOKEN 으로 저장하세요.

const SHEET_NAME = '운동기록';
const HEADERS = ['id', '날짜', '운동 종류', '세트', '횟수', '무게(kg)', '시간(분)', '거리(km)', '메모'];
const FIELDS = ['type', 'sets', 'reps', 'weight', 'minutes', 'distance', 'memo'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // 날짜를 문자열 그대로 유지
  }
  return sh;
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return reply_({ ok: false, error: 'bad request' }); }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || req.token !== token) return reply_({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    if (req.action === 'upsert') {
      const en = req.entry;
      const row = [en.id, req.date].concat(FIELDS.map(f => en[f] === undefined ? '' : en[f]));
      const at = findRow_(sh, en.id);
      if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]);
      else {
        sh.appendRow(row);
        sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(req.date);
      }
      return reply_({ ok: true });
    }
    if (req.action === 'delete') {
      const at = findRow_(sh, req.id);
      if (at > 0) sh.deleteRow(at);
      return reply_({ ok: true });
    }
    if (req.action === 'list') {
      const last = sh.getLastRow();
      const logs = {};
      if (last >= 2) {
        sh.getRange(2, 1, last - 1, HEADERS.length).getValues().forEach(r => {
          if (!r[0]) return;
          const date = r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1]);
          const en = { id: String(r[0]) };
          FIELDS.forEach((f, i) => en[f] = r[i + 2] === '' ? '' : String(r[i + 2]));
          (logs[date] = logs[date] || []).push(en);
        });
      }
      return reply_({ ok: true, logs: logs });
    }
    return reply_({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}
