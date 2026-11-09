/**
 * 차계부 API (Google Apps Script)
 * script.google.com 의 새 프로젝트에 붙여 넣고 SPREADSHEET_ID 를 지정한 뒤
 * [배포 > 새 배포 > 웹 앱] 으로 배포합니다.  (실행 사용자: 나 / 액세스: 모든 사용자)
 *
 *  GET  ?action=all                  → 모든 데이터 시트 내용을 JSON으로 반환
 *  POST {action:'append', sheet, row}          → 행 추가
 *  POST {action:'update', sheet, id, row}      → ID로 찾아 일부 열 수정
 *  POST {action:'delete', sheet, id}           → ID로 찾아 행 삭제
 *  POST {action:'batch',  ops:[...]}           → 위 동작 여러 개를 한 번에 처리
 *  (POST 는 Content-Type: text/plain 으로 보냅니다 — 브라우저 CORS 사전요청 회피)
 */

// 데이터가 들어 있는 구글 스프레드시트의 ID (주소창 .../d/<여기>/edit)
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
const TZ = 'Asia/Seoul';

// 웹앱이 읽고 쓸 수 있는 시트와 각 시트의 ID 열 이름 (S_항목현황은 수식 시트라 제외)
const SHEETS = {
  'M_차량': 'ID',
  'M_정비항목': '항목코드',
  'M_업체': 'ID',
  'R_정비이력': 'ID',
  'R_주행거리': 'ID',
  'R_세금보험': 'ID',
  'R_사고': 'ID',
  'P_로드맵': 'ID'
};

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

let tzCache_ = null;
function tz_() {
  if (!tzCache_) tzCache_ = ss_().getSpreadsheetTimeZone() || TZ;
  return tzCache_;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    if (action === 'all') return json_({ ok: true, data: readAll_() });
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const req = JSON.parse(e.postData.contents);
    const ops = req.action === 'batch' ? req.ops : [req];
    ops.forEach(applyOp_);
    SpreadsheetApp.flush();
    return json_({ ok: true, data: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- 읽기 ---------- */
function readAll_() {
  const out = {};
  Object.keys(SHEETS).forEach(function (name) { out[name] = readSheet_(name); });
  return out;
}

function readSheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('시트가 없습니다: ' + name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  const idCol = headers.indexOf(SHEETS[name]);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (idCol >= 0 && (values[i][idCol] === '' || values[i][idCol] === null)) continue;
    const o = {};
    headers.forEach(function (h, j) { o[h] = norm_(values[i][j]); });
    rows.push(o);
  }
  return rows;
}

function norm_(v) {
  // 날짜 셀은 '시트에 설정된 시간대' 기준 자정이므로 같은 시간대로 변환해야 날짜가 어긋나지 않음
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  if (v === '' || v === null || v === undefined) return null;
  return v;
}

/* ---------- 쓰기 ---------- */
function applyOp_(op) {
  if (!SHEETS.hasOwnProperty(op.sheet)) throw new Error('허용되지 않은 시트: ' + op.sheet);
  const sh = ss_().getSheetByName(op.sheet);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idName = SHEETS[op.sheet];
  const idCol = headers.indexOf(idName);

  if (op.action === 'append') {
    const arr = headers.map(function (h) { return clean_(op.row[h]); });
    const target = sh.getLastRow() + 1;
    sh.getRange(target, 1, 1, arr.length).setValues([arr]);
    return;
  }

  const row = findRow_(sh, idCol, op.id);
  if (row < 0) throw new Error('해당 ID를 찾을 수 없습니다: ' + op.id);

  if (op.action === 'update') {
    Object.keys(op.row || {}).forEach(function (h) {
      const c = headers.indexOf(h);
      if (c < 0 || h === idName || h === '현재주행거리') return;
      sh.getRange(row, c + 1).setValue(clean_(op.row[h]));
    });
  } else if (op.action === 'delete') {
    sh.deleteRow(row);
  } else {
    throw new Error('unknown action: ' + op.action);
  }
}

function findRow_(sh, idCol, id) {
  if (idCol < 0) return -1;
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, idCol + 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

// 빈 값은 공란, 수식으로 해석될 수 있는 문자열은 막기
function clean_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && /^[=+@]/.test(v)) return "'" + v;
  return v;
}

/** 배포 전 동작 확인용: 편집기에서 실행하면 로그에 시트별 행 수가 표시됩니다. */
function testRead() {
  const d = readAll_();
  Object.keys(d).forEach(function (k) { Logger.log(k + ': ' + d[k].length + '행'); });
}
