/**
 * 단체티 사이즈 조사 — Google Apps Script 백엔드
 * 구글 시트 [확장 프로그램] → [Apps Script]에 이 코드를 통째로 붙여넣고 웹앱으로 배포하세요.
 */

// ▼ 선생님용 관리자 비밀번호 (꼭 바꾸세요)
const ADMIN_KEY = '1234';

const SHEET_NAME = '응답';
const HEADER = ['학년', '반', '번호', '이름', '사이즈', '제출시각', '수정횟수'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'submit') return out(submit(d));
    if (d.action === 'list')   return out(d.key === ADMIN_KEY ? { ok: true, rows: list() } : { ok: false, error: 'auth' });
    if (d.action === 'delete') return out(d.key === ADMIN_KEY ? remove(d) : { ok: false, error: 'auth' });
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() {
  return out({ ok: true, msg: '단체티 사이즈 조사 API가 동작 중입니다.' });
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADER.length).setFontWeight('bold').setBackground('#e8eefe');
  }
  return sh;
}

function submit(d) {
  const grade = parseInt(d.grade, 10), cls = parseInt(d.cls, 10), no = parseInt(d.no, 10);
  const name = String(d.name || '').trim().slice(0, 20);
  const size = String(d.size || '').trim().slice(0, 6);
  if (!grade || !cls || !no || name.length < 2 || !size) return { ok: false, error: 'invalid' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const values = sh.getDataRange().getValues();
    const now = new Date();
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      if (r[0] == grade && r[1] == cls && r[2] == no) {
        sh.getRange(i + 1, 1, 1, 7).setValues([[grade, cls, no, name, size, now, (Number(r[6]) || 0) + 1]]);
        return { ok: true, updated: true };
      }
    }
    sh.appendRow([grade, cls, no, name, size, now, 0]);
    return { ok: true, updated: false };
  } finally {
    lock.releaseLock();
  }
}

function list() {
  const values = sheet().getDataRange().getValues();
  return values.slice(1).filter(r => r[0] !== '').map(r => ({
    grade: Number(r[0]), cls: Number(r[1]), no: Number(r[2]),
    name: String(r[3]), size: String(r[4]),
    time: r[5] instanceof Date ? r[5].toISOString() : String(r[5])
  }));
}

function remove(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const values = sh.getDataRange().getValues();
    for (let i = values.length - 1; i >= 1; i--) {
      const r = values[i];
      if (r[0] == d.grade && r[1] == d.cls && r[2] == d.no) sh.deleteRow(i + 1);
    }
    return { ok: true, rows: list() };
  } finally {
    lock.releaseLock();
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
