/**
 * 단체티 사이즈 조사 — Google Apps Script 백엔드 (속도 개선판)
 * 구글 시트 [확장 프로그램] → [Apps Script]에 이 코드를 통째로 붙여넣고
 * [배포] → [배포 관리] → 연필(편집) → 버전: "새 버전" → 배포  (URL은 그대로 유지됨)
 *
 * 속도 개선 내용
 *  - 제출: 시트 전체를 읽고 잠금(Lock)을 거는 대신, 맨 아래에 한 줄만 추가 → 여러 학생이 동시에 내도 줄 서지 않음
 *  - 같은 학생이 여러 번 내면 시트에는 여러 줄이 남고, 집계 때 "가장 마지막 제출"만 셉니다
 *  - 집계 결과를 30초 동안 캐시 → 선생님 화면 새로고침이 빨라짐 (새 제출이 들어오면 캐시 자동 삭제)
 *  - 팀 배정(두레/자연)은 '팀배정' 탭에 따로 저장 → 학생이 사이즈를 다시 내도 팀은 유지됨
 */

// ▼ 선생님용 관리자 비밀번호 (꼭 바꾸세요)
const ADMIN_KEY = '1234';

const SHEET_NAME = '응답';
const HEADER = ['학년', '반', '번호', '이름', '사이즈', '제출시각'];
const TEAM_SHEET = '팀배정';
const CACHE_KEY = 'list_v3';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'submit') return out(submit(d));
    if (d.action === 'list')   return out(d.key === ADMIN_KEY ? { ok: true, rows: list() } : { ok: false, error: 'auth' });
    if (d.action === 'setTeams') return out(d.key === ADMIN_KEY ? setTeams(d.teams || {}) : { ok: false, error: 'auth' });
    if (d.action === 'delete') return out(d.key === ADMIN_KEY ? remove(d) : { ok: false, error: 'auth' });
    return out({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

// 페이지를 열 때 보내는 '예열' 요청 — 아무 일도 하지 않고 바로 응답
function doGet() {
  return out({ ok: true });
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

  // 이미 낸 적 있는지는 캐시된 목록으로만 확인 (시트 전체를 다시 읽지 않음)
  let updated = false;
  const cached = CacheService.getScriptCache().get(CACHE_KEY);
  if (cached) updated = JSON.parse(cached).some(r => r.grade === grade && r.cls === cls && r.no === no);

  sheet().appendRow([grade, cls, no, name, size, new Date()]);   // appendRow는 동시에 여러 명이 불러도 안전
  CacheService.getScriptCache().remove(CACHE_KEY);
  return { ok: true, updated: updated };
}

// 학생별(학년-반-번호) 마지막 제출만 남겨서 돌려줌
function list() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);

  const sh = sheet();
  const last = sh.getLastRow();
  const map = {};
  if (last > 1) {
    const values = sh.getRange(2, 1, last - 1, 6).getValues();
    values.forEach(r => {
      if (r[0] === '') return;
      const key = r[0] + '-' + r[1] + '-' + r[2];
      map[key] = {                                  // 아래쪽(나중) 줄이 위쪽 줄을 덮어씀
        grade: Number(r[0]), cls: Number(r[1]), no: Number(r[2]),
        name: String(r[3]), size: String(r[4]),
        time: r[5] instanceof Date ? r[5].toISOString() : String(r[5])
      };
    });
  }
  const teams = readTeams();
  const rows = Object.keys(map).map(k => { map[k].team = teams[k] || ''; return map[k]; });
  try { cache.put(CACHE_KEY, JSON.stringify(rows), 30); } catch (e) {}   // 30초 캐시
  return rows;
}

/* ===== 팀 배정 ===== */
function teamSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TEAM_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TEAM_SHEET);
    sh.appendRow(['학년', '반', '번호', '팀']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e7f6ef');
  }
  return sh;
}

function readTeams() {
  const sh = teamSheet();
  const last = sh.getLastRow();
  const map = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 4).getValues().forEach(r => {
      if (r[0] !== '') map[r[0] + '-' + r[1] + '-' + r[2]] = String(r[3] || '');
    });
  }
  return map;
}

// teams = { "1-1-3": "두레", "2-1-5": "자연", "3-1-2": "" (해제) }
function setTeams(teams) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = teamSheet();
    const merged = readTeams();
    Object.keys(teams).forEach(k => { merged[k] = String(teams[k] || '').slice(0, 10); });
    const values = Object.keys(merged)
      .filter(k => merged[k])
      .map(k => { const p = k.split('-').map(Number); return [p[0], p[1], p[2], merged[k]]; })
      .sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]);
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
    if (values.length) sh.getRange(2, 1, values.length, 4).setValues(values);   // 한 번에 쓰기
    CacheService.getScriptCache().remove(CACHE_KEY);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function remove(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const last = sh.getLastRow();
    if (last > 1) {
      const values = sh.getRange(2, 1, last - 1, 3).getValues();
      for (let i = values.length - 1; i >= 0; i--) {
        const r = values[i];
        if (r[0] == d.grade && r[1] == d.cls && r[2] == d.no) sh.deleteRow(i + 2);
      }
    }
    CacheService.getScriptCache().remove(CACHE_KEY);
    return { ok: true, rows: list() };
  } finally {
    lock.releaseLock();
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
