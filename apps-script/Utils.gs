/**
 * 공통 도구: 시트 읽기, 해시, 토큰, 잠금, 로그
 */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    throw new Error('"' + name + '" 탭이 없습니다. 메뉴에서 "초기 설정"을 먼저 실행해 주세요.');
  }
  return sh;
}

function alert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (err) {
    console.log(msg);
  }
}

function confirm_(msg) {
  var ui = SpreadsheetApp.getUi();
  return ui.alert('확인', msg, ui.ButtonSet.YES_NO) === ui.Button.YES;
}

/* ---------- 이름, 전화번호, 해시 ---------- */

function normName_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '');
}

/** 숫자만 남기고 010 형식으로 맞춥니다. (+82 10..., 앞자리 0이 빠진 10... 도 처리) */
function normPhone_(s) {
  var d = String(s == null ? '' : s).replace(/\D/g, '');
  if (d.indexOf('82') === 0 && d.length >= 11) d = '0' + d.substring(2);
  if (d.length === 10 && d.indexOf('10') === 0) d = '0' + d;
  return d;
}

function getSalt_() {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('SALT');
  if (!salt) {
    salt = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SALT', salt);
  }
  return salt;
}

function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    out += (b < 16 ? '0' : '') + b.toString(16);
  }
  return out;
}

function hashPhone_(phone) {
  return sha256Hex_(getSalt_() + normPhone_(phone));
}

function isHash_(s) {
  return /^[0-9a-f]{64}$/.test(String(s));
}

/** 수식으로 해석될 수 있는 앞 글자를 떼어 냅니다. */
function cleanText_(s, maxLen) {
  var t = String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim();
  t = t.replace(/^[=+\-@\s]+/, '');
  return t.substring(0, maxLen || 200);
}

/* ---------- 날짜 ---------- */

function fmt_(date, pattern) {
  return Utilities.formatDate(date, TZ, pattern);
}

/** 'yyyy-MM-dd' 문자열의 요일 */
function dowOf_(dateStr) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(dateStr));
  if (!m) return '';
  var d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  return ['일', '월', '화', '수', '목', '금', '토'][d.getUTCDay()];
}

/** '2026-12-29' → '12/29(화)' */
function shortDate_(slot) {
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(slot.date));
  var base = m ? Number(m[2]) + '/' + Number(m[3]) : String(slot.date);
  return base + (slot.dow ? '(' + slot.dow + ')' : '');
}

/** 팀 열의 표시 이름. A, B는 "A팀", 그 밖의 값(함께할 사람 이름 등)은 그대로 씁니다. */
function teamLabel_(team) {
  if (!team || team === '공통') return '공통';
  return /^[A-Za-z]$/.test(team) ? team + '팀' : team;
}

function slotLabel_(slot) {
  var item = String(slot.item).replace('봉사', '').replace('식사', '');
  var team = slot.team && slot.team !== '공통' ? ' ' + slot.team : '';
  var time = isService_(slot) && slot.time ? ' ' + slot.time : '';
  return shortDate_(slot) + ' ' + item + time + team;
}

/** '9:45', '13:45 ①' 같은 글자에서 분 단위 시각을 뽑습니다. 없으면 -1. */
function timeMinutes_(text) {
  var m = /(\d{1,2}):(\d{2})/.exec(String(text));
  return m ? Number(m[1]) * 60 + Number(m[2]) : -1;
}

function isService_(slot) {
  return String(slot.item).indexOf('봉사') >= 0;
}

function isDinner_(slot) {
  return String(slot.item).indexOf('식사') >= 0;
}

/* ---------- 시트 읽기 ---------- */

/** A열 기준으로 마지막 자료 행을 찾습니다. (체크박스 때문에 getLastRow가 커지는 것을 피함) */
function lastDataRow_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var col = sh.getRange(1, 1, last, 1).getValues();
  for (var i = col.length - 1; i >= 1; i--) {
    if (String(col[i][0]).trim() !== '') return i + 1;
  }
  return 1;
}

function appendRows_(sh, rows) {
  if (!rows.length) return;
  var start = lastDataRow_(sh) + 1;
  var need = start + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function loadSlots_() {
  var sh = sheet_(TAB.SCHEDULE);
  var last = lastDataRow_(sh);
  if (last < 2) return [];
  var range = sh.getRange(2, 1, last - 1, HEAD.SCHEDULE.length);
  var v = range.getValues();
  var d = range.getDisplayValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var id = String(v[i][0]).trim();
    if (!id) continue;
    var date = v[i][1] instanceof Date ? fmt_(v[i][1], 'yyyy-MM-dd') : String(d[i][1]).trim();
    out.push({
      id: id,
      date: date,
      dow: String(d[i][2]).trim() || dowOf_(date),
      item: String(d[i][3]).trim(),
      team: String(d[i][4]).trim() || '공통',
      time: String(d[i][5]).trim(),
      place: String(d[i][6]).trim(),
      need: Number(v[i][7]) || 0,
      open: v[i][8] === true,
      prefer: String(d[i][9]).trim(),
      order: i
    });
  }
  out.sort(compareSlots_);
  return out;
}

function compareSlots_(a, b) {
  if (a.date !== b.date) return a.date < b.date ? -1 : 1;
  var ia = ITEMS.indexOf(a.item), ib = ITEMS.indexOf(b.item);
  if (ia < 0) ia = 9;
  if (ib < 0) ib = 9;
  if (ia !== ib) return ia - ib;
  var ma = timeMinutes_(a.time), mb = timeMinutes_(b.time);
  if (ma !== mb) return ma - mb;
  if (a.time !== b.time) return a.time < b.time ? -1 : 1;
  return a.order - b.order;  // 같은 시간이면 일정 탭에 적은 순서
}

/** 신청 탭 전체. row는 시트의 실제 행 번호입니다. */
function loadApps_() {
  var sh = sheet_(TAB.APPLY);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, HEAD.APPLY.length).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var slotId = String(v[i][0]).trim();
    var name = String(v[i][1]).trim();
    if (!slotId || !name) continue;
    out.push({
      row: i + 2,
      slotId: slotId,
      name: name,
      time: v[i][2],
      status: String(v[i][3]).trim() || ST.APPLY,
      car: String(v[i][4]).trim(),
      meal: String(v[i][5]).trim(),
      teams: v[i][6] === '' ? '' : Number(v[i][6]) || '',
      memo: String(v[i][7])
    });
  }
  return out;
}

function loadRoster_() {
  var sh = sheet_(TAB.ROSTER);
  var last = lastDataRow_(sh);
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, HEAD.ROSTER.length).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var name = String(v[i][0]).trim();
    if (!name) continue;
    out.push({
      row: i + 2,
      name: name,
      key: normName_(name),
      hash: String(v[i][1]).trim(),
      gender: String(v[i][2]).trim()
    });
  }
  return out;
}

function loadConfig_() {
  var v = sheet_(TAB.CONFIG).getRange(2, 2, 2, 1).getValues();
  var deadline = v[0][0] instanceof Date ? v[0][0] : null;
  return {
    deadline: deadline,
    closed: deadline ? new Date().getTime() > deadline.getTime() : false,
    published: v[1][0] === true
  };
}

function loadAll_() {
  return { slots: loadSlots_(), apps: loadApps_(), roster: loadRoster_() };
}

function appKey_(slotId, name) {
  return slotId + '|' + name;
}

/* ---------- 잠금 ---------- */

function apiError_(code, message) {
  var err = new Error(message);
  err.apiCode = code;
  return err;
}

/** 스크립트 잠금을 걸고 fn을 실행합니다. 최대 10초 기다립니다. */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw apiError_('BUSY', '지금 접속이 많습니다. 잠시 후 다시 눌러 주세요.');
  }
  try {
    var result = fn();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 로그 ---------- */

function log_(name, action, slotId) {
  sheet_(TAB.LOG).appendRow([new Date(), name, action, slotId || '']);
}

function logMany_(rows) {
  if (!rows.length) return;
  var sh = sheet_(TAB.LOG);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}

/* ---------- 트리거 ---------- */

function deleteTriggers_(handler) {
  var triggers = ScriptApp.getProjectTriggers();
  var n = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(triggers[i]);
      n++;
    }
  }
  return n;
}
