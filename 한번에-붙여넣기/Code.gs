// 순회 방문 신청 - 한 번에 붙여 넣는 파일
// apps-script 폴더의 파일 8개를 하나로 합친 것입니다. 스크립트 편집기의 Code.gs에 통째로 붙여 넣으세요.
// (apps-script 폴더의 파일을 고치면 이 파일도 다시 만들어야 합니다.)

// ==================== Code.gs ====================
/**
 * 순회 방문 신청 - 공통 상수와 시트 메뉴
 *
 * 파일 구성
 *   Code.gs       상수, 시트 메뉴
 *   Utils.gs      해시, 토큰, 잠금, 로그, 시트 읽기
 *   Setup.gs      초기 설정, 전화번호 해시 변환
 *   Api.gs        웹앱 API (login / getData / toggle / updateExtra)
 *   Status.gs     신청현황, 전도인별 탭 그리기와 자동 갱신
 *   Sync.gs       배정 체크박스 동기화, 대기 표시
 *   Recommend.gs  자동 추천 배정
 *   TestData.gs   테스트 데이터 넣기 / 지우기
 */

var TZ = 'Asia/Seoul';

var TAB = {
  ROSTER: '명단',
  SCHEDULE: '일정',
  APPLY: '신청',
  CONFIG: '설정',
  STATUS: '신청현황',
  RECO: '추천결과',
  PERSON: '전도인별',
  LOG: '로그'
};

var HEAD = {
  ROSTER: ['이름', '전화번호해시', '성별'],
  SCHEDULE: ['칸ID', '날짜', '요일', '항목', '팀', '시간', '장소', '필요인원', '열림여부', '추천성별'],
  APPLY: ['칸ID', '이름', '신청시각', '상태', '차량제공', '식사방식', '대접가능팀수', '메모'],
  PERSON: ['이름', '성별', '신청한 칸', '배정된 칸', '신청 수', '배정 수'],
  LOG: ['시각', '이름', '동작', '칸ID']
};

var ST = { APPLY: '신청', ASSIGN: '배정', WAIT: '대기', CANCEL: '취소' };

var ITEMS = ['오전봉사', '점심식사', '오후봉사', '저녁식사'];

// 신청현황 / 추천결과 탭의 열 위치 (1부터 시작)
var VIEW = {
  CHECK: 1, NAME: 2, GENDER: 3, CAR: 4, MEAL: 5, TEAMS: 6, MEMO: 7,
  TOTAL_APPLY: 8, TOTAL_ASSIGN: 9, TIME: 10, KEY_SLOT: 11, KEY_NAME: 12,
  WIDTH: 12,
  FIRST_ROW: 3,          // 1행: 갱신 시각, 2행: 열 이름
  HEADER_MARK: '#H'      // 블록 머리줄 표시 (숨김 열)
};

var COLOR = {
  HEAD_NORMAL: '#cfe2f3',
  HEAD_SHORT: '#f9cb9c',   // 신청자가 필요 인원보다 적음
  HEAD_FULL: '#b6d7a8',    // 배정이 필요 인원만큼 참
  LABEL: '#eeeeee',
  RECO: '#fff2cc'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('순회방문 관리')
    .addItem('초기 설정', 'setup')
    .addItem('전화번호 해시 변환', 'convertPhones')
    .addSeparator()
    .addItem('신청현황 새로고침', 'refreshStatus')
    .addItem('자동 갱신 켜기', 'enableAutoRefresh')
    .addItem('자동 갱신 끄기', 'disableAutoRefresh')
    .addItem('배정 동기화 트리거 설치', 'installSyncTrigger')
    .addSeparator()
    .addItem('자동 추천 배정 (미리보기 생성)', 'recommend')
    .addItem('추천 결과 반영', 'applyRecommendation')
    .addItem('체크 안 된 신청을 모두 대기로 표시', 'markWaiting')
    .addSeparator()
    .addItem('테스트 데이터 넣기', 'insertTestData')
    .addItem('테스트 데이터 지우기', 'removeTestData')
    .addToUi();
}

// ==================== Utils.gs ====================
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

// ==================== Setup.gs ====================
/**
 * 초기 설정과 전화번호 해시 변환
 */

/** 메뉴: 초기 설정. 여러 번 실행해도 입력해 둔 자료는 지우지 않습니다. */
function setup() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  getSalt_();

  var dv = SpreadsheetApp.newDataValidation;
  var rows = 500;

  // 1. 명단
  var roster = ensureSheet_(TAB.ROSTER, HEAD.ROSTER, [120, 260, 80]);
  roster.getRange(2, 2, rows, 1).setNumberFormat('@');
  roster.getRange(2, 3, rows, 1).setDataValidation(
    dv().requireValueInList(['형제', '자매'], true).setAllowInvalid(false).build());
  roster.getRange(1, 2).setNote(
    '전화번호를 그대로 입력한 뒤 메뉴에서 "전화번호 해시 변환"을 누르면 암호화된 값으로 바뀝니다.');

  // 2. 일정
  var sched = ensureSheet_(TAB.SCHEDULE, HEAD.SCHEDULE, [150, 110, 50, 90, 90, 90, 200, 80, 80, 80]);
  sched.getRange(2, 1, rows, 1).setNumberFormat('@');
  sched.getRange(2, 2, rows, 1).setNumberFormat('yyyy-mm-dd');
  sched.getRange(2, 4, rows, 1).setDataValidation(
    dv().requireValueInList(ITEMS, true).setAllowInvalid(true).build());
  sched.getRange(2, 5, rows, 1).clearDataValidations();
  sched.getRange(2, 10, rows, 1).setDataValidation(
    dv().requireValueInList(['형제', '자매'], true).setAllowInvalid(false).build());
  sched.getRange(2, 6, rows, 1).setNumberFormat('@');
  sched.getRange(2, 9, rows, 1).setDataValidation(dv().requireCheckbox().build());
  sched.getRange(1, 3).setNote('비워 두면 날짜에서 자동으로 계산합니다.');
  sched.getRange(1, 5).setNote('A, B 또는 함께할 사람의 이름을 적습니다. 모두 함께하는 칸은 "공통"이라고 적습니다.');
  sched.getRange(1, 9).setNote('체크한 칸만 전도인이 신청할 수 있습니다.');
  sched.getRange(1, 10).setNote('적어 두면 자동 추천이 그 성별의 신청자만 추천합니다. 비워 두면 가리지 않습니다.');

  // 3. 신청 (모든 상태의 원본)
  var apply = ensureSheet_(TAB.APPLY, HEAD.APPLY, [130, 120, 150, 70, 80, 80, 100, 260]);
  apply.getRange(2, 1, rows, 2).setNumberFormat('@');
  apply.getRange(2, 3, rows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  apply.getRange(2, 4, rows, 1).setDataValidation(
    dv().requireValueInList([ST.APPLY, ST.ASSIGN, ST.WAIT, ST.CANCEL], true).setAllowInvalid(false).build());
  apply.getRange(2, 8, rows, 1).setNumberFormat('@');

  // 4. 설정
  var conf = ensureSheet_(TAB.CONFIG, ['항목', '값', '설명'], [140, 180, 420]);
  conf.getRange(2, 1, 2, 1).setValues([['신청마감일시'], ['결과공개여부']]);
  conf.getRange(2, 3, 2, 1).setValues([
    ['예: 2026-12-20 21:00  (이 시각이 지나면 신청과 취소가 막힙니다. 비워 두면 마감이 없습니다.)'],
    ['체크하면 전도인 화면에 각자의 배정 결과가 보입니다.']
  ]);
  conf.getRange(2, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  var pub = conf.getRange(3, 2);
  pub.setDataValidation(dv().requireCheckbox().build());
  if (pub.getValue() !== true) pub.setValue(false);

  // 5~7. 보기용 탭
  ensureSheet_(TAB.STATUS, null, null);
  ensureSheet_(TAB.RECO, null, null);
  ensureSheet_(TAB.PERSON, HEAD.PERSON, [120, 60, 420, 420, 70, 70]);

  // 8. 로그
  var log = ensureSheet_(TAB.LOG, HEAD.LOG, [150, 120, 120, 130]);
  log.getRange(2, 1, rows, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  refreshAll_();
  alert_('초기 설정이 끝났습니다.\n\n' +
    '다음 순서로 진행해 주세요.\n' +
    '1. 명단 탭에 이름, 전화번호, 성별 입력 → "전화번호 해시 변환"\n' +
    '2. 일정 탭에 계획표 입력\n' +
    '3. 설정 탭에 신청마감일시 입력\n' +
    '4. "자동 갱신 켜기"와 "배정 동기화 트리거 설치" 실행');
}

function ensureSheet_(name, headers, widths) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, ss.getNumSheets());
  if (headers) {
    sh.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground(COLOR.LABEL);
    sh.setFrozenRows(1);
  }
  if (widths) {
    for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
  }
  return sh;
}

/** 메뉴: 전화번호 해시 변환. 명단 탭 B열의 전화번호 원문을 해시로 바꿉니다. */
function convertPhones() {
  var result = withLock_(function () {
    var sh = sheet_(TAB.ROSTER);
    var roster = loadRoster_();
    var converted = 0, already = 0, empty = [], invalid = [];
    var seen = {}, dup = {};

    for (var i = 0; i < roster.length; i++) {
      var p = roster[i];
      if (seen[p.key]) dup[p.name] = true;
      seen[p.key] = true;

      if (!p.hash) { empty.push(p.name); continue; }
      if (isHash_(p.hash)) { already++; continue; }
      var digits = normPhone_(p.hash);
      if (digits.length < 10 || digits.length > 11) { invalid.push(p.name); continue; }
      sh.getRange(p.row, 2).setNumberFormat('@').setValue(hashPhone_(digits));
      converted++;
    }
    return { converted: converted, already: already, empty: empty, invalid: invalid, dup: Object.keys(dup) };
  });

  var msg = '변환 ' + result.converted + '명, 이미 변환됨 ' + result.already + '명';
  if (result.invalid.length) msg += '\n\n전화번호 형식이 이상한 분: ' + result.invalid.join(', ');
  if (result.empty.length) msg += '\n\n전화번호가 비어 있는 분: ' + result.empty.join(', ');
  if (result.dup.length) {
    msg += '\n\n이름이 겹치는 분: ' + result.dup.join(', ') +
      '\n→ "김철수(대)", "김철수(소)"처럼 서로 다르게 고쳐 주세요.';
  }
  alert_(msg);
}

// ==================== Api.gs ====================
/**
 * 웹앱 API. 화면(GitHub Pages)이 POST로 호출합니다.
 * 본문은 text/plain으로 오는 JSON이며, action 필드로 동작을 구분합니다.
 */

var TOKEN_SECONDS = 21600;      // 6시간
var LOGIN_MAX_FAIL = 5;
var LOGIN_BLOCK_SECONDS = 600;  // 10분

function doGet() {
  return json_({ ok: true, message: '순회 방문 신청 API가 동작 중입니다.' });
}

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    switch (req.action) {
      case 'login': out = apiLogin_(req); break;
      case 'getData': out = apiGetData_(req); break;
      case 'toggle': out = apiToggle_(req); break;
      case 'updateExtra': out = apiUpdateExtra_(req); break;
      default: throw apiError_('BAD_REQUEST', '알 수 없는 요청입니다.');
    }
  } catch (err) {
    if (err && err.apiCode) {
      out = { ok: false, code: err.apiCode, message: err.message };
    } else {
      console.error(err && err.stack ? err.stack : err);
      out = { ok: false, code: 'ERROR', message: '문제가 생겼습니다. 잠시 후 다시 시도해 주세요.' };
    }
  }
  return json_(out);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 로그인과 토큰 ---------- */

function apiLogin_(req) {
  var nameKey = normName_(req.name).substring(0, 30);
  var phone = normPhone_(String(req.phone || '').substring(0, 30));
  if (!nameKey || !phone) {
    throw apiError_('BAD_REQUEST', '이름과 휴대폰 번호를 모두 입력해 주세요.');
  }

  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + sha256Hex_(nameKey).substring(0, 40);
  var fails = Number(cache.get(failKey)) || 0;
  if (fails >= LOGIN_MAX_FAIL) {
    throw apiError_('BLOCKED', '여러 번 틀려서 잠시 막혔습니다. 10분 뒤에 다시 시도해 주세요.');
  }

  var hash = hashPhone_(phone);
  var roster = loadRoster_();
  var found = null;
  for (var i = 0; i < roster.length; i++) {
    if (roster[i].key === nameKey && roster[i].hash === hash) { found = roster[i]; break; }
  }

  if (!found) {
    fails++;
    cache.put(failKey, String(fails), LOGIN_BLOCK_SECONDS);
    if (fails >= LOGIN_MAX_FAIL) {
      throw apiError_('BLOCKED', '여러 번 틀려서 잠시 막혔습니다. 10분 뒤에 다시 시도해 주세요.');
    }
    throw apiError_('NOT_FOUND',
      '이름이나 휴대폰 번호가 명단과 다릅니다. 다시 확인해 주세요. (' + fails + '/' + LOGIN_MAX_FAIL + ')');
  }

  cache.remove(failKey);
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  cache.put('tok_' + token, found.name, TOKEN_SECONDS);
  log_(found.name, '로그인', '');
  return { ok: true, token: token, name: found.name };
}

/** 토큰으로 본인을 확인하고 이름을 돌려줍니다. 화면이 보낸 이름은 쓰지 않습니다. */
function auth_(token) {
  token = String(token || '');
  var name = /^[0-9a-f]{64}$/.test(token) ? CacheService.getScriptCache().get('tok_' + token) : null;
  if (!name) {
    throw apiError_('AUTH', '로그인 시간이 지났습니다. 다시 로그인해 주세요.');
  }
  return name;
}

/* ---------- getData ---------- */

function apiGetData_(req) {
  var name = auth_(req.token);
  var cfg = loadConfig_();
  var slots = loadSlots_();
  var apps = loadApps_();

  var counts = {}, mine = {};
  for (var i = 0; i < apps.length; i++) {
    var a = apps[i];
    if (a.status === ST.CANCEL) continue;
    counts[a.slotId] = (counts[a.slotId] || 0) + 1;
    if (a.name === name) {
      mine[a.slotId] = {
        status: publicStatus_(a.status, cfg.published),
        car: a.car !== '',
        meal: a.meal,
        teams: a.teams,
        memo: a.memo
      };
    }
  }

  var outSlots = [];
  for (var j = 0; j < slots.length; j++) {
    var s = slots[j];
    outSlots.push({
      id: s.id, date: s.date, dow: s.dow, item: s.item, team: s.team,
      time: s.time, place: s.place, open: s.open, count: counts[s.id] || 0
    });
  }

  return {
    ok: true,
    name: name,
    slots: outSlots,
    mine: mine,
    deadlineText: cfg.deadline ? deadlineText_(cfg.deadline) : '',
    closed: cfg.closed,
    published: cfg.published
  };
}

/** 결과 공개 전에는 배정 여부가 드러나지 않게 모두 "신청"으로 보냅니다. */
function publicStatus_(status, published) {
  if (!published) return ST.APPLY;
  return status === ST.ASSIGN ? ST.ASSIGN : ST.WAIT;
}

function deadlineText_(d) {
  var h = Number(fmt_(d, 'H'));
  var ampm = h < 12 ? '오전' : '오후';
  var h12 = h % 12 === 0 ? 12 : h % 12;
  var dow = dowOf_(fmt_(d, 'yyyy-MM-dd'));
  return fmt_(d, 'M월 d일') + '(' + dow + ') ' + ampm + ' ' + h12 + ':' + fmt_(d, 'mm');
}

/* ---------- toggle / updateExtra ---------- */

/** 마감, 결과 공개, 칸 열림 여부를 확인하고 칸 정보를 돌려줍니다. */
function checkOpen_(slotId) {
  var cfg = loadConfig_();
  if (cfg.closed || cfg.published) {
    throw apiError_('CLOSED', '신청이 마감되었습니다. 바꾸실 내용이 있으면 장로에게 연락해 주세요.');
  }
  var slots = loadSlots_();
  for (var i = 0; i < slots.length; i++) {
    if (slots[i].id === slotId) {
      if (!slots[i].open) throw apiError_('CLOSED', '이 칸은 지금 신청을 받지 않습니다.');
      return slots[i];
    }
  }
  throw apiError_('CLOSED', '없는 칸입니다. 화면을 새로 고쳐 주세요.');
}

function cleanExtra_(extra, slot) {
  extra = extra || {};
  var out = { car: '', meal: '', teams: '', memo: cleanText_(extra.memo, 200) };
  if (isService_(slot)) {
    out.car = extra.car === true ? 'O' : '';
  }
  if (isDinner_(slot)) {
    out.meal = extra.meal === '집' || extra.meal === '외식' ? extra.meal : '';
    if (slot.team === '공통') {
      out.teams = Number(extra.teams) === 2 ? 2 : 1;
    }
  }
  return out;
}

function findApp_(apps, slotId, name) {
  for (var i = 0; i < apps.length; i++) {
    if (apps[i].slotId === slotId && apps[i].name === name) return apps[i];
  }
  return null;
}

function countApplied_(apps, slotId) {
  var n = 0;
  for (var i = 0; i < apps.length; i++) {
    if (apps[i].slotId === slotId && apps[i].status !== ST.CANCEL) n++;
  }
  return n;
}

/**
 * 신청 또는 취소.
 * want: 'on'(신청) / 'off'(취소) 를 보내면 그 상태로 맞추고, 없으면 현재 상태를 뒤집습니다.
 */
function apiToggle_(req) {
  var name = auth_(req.token);
  var slotId = String(req.slotId || '');
  var slot = checkOpen_(slotId);
  var extra = cleanExtra_(req.extra, slot);
  var want = req.want === 'on' ? true : (req.want === 'off' ? false : null);

  return withLock_(function () {
    var sh = sheet_(TAB.APPLY);
    var apps = loadApps_();
    var app = findApp_(apps, slotId, name);
    var applied;

    if (!app) {
      if (want === false) {
        applied = false;
      } else {
        sh.appendRow([slotId, name, new Date(), ST.APPLY, extra.car, extra.meal, extra.teams, extra.memo]);
        log_(name, '신청', slotId);
        applied = true;
      }
    } else if (app.status === ST.ASSIGN) {
      if (want === true) {
        applied = true;
      } else {
        throw apiError_('ASSIGNED', '이미 배정된 칸입니다. 취소하시려면 장로에게 연락해 주세요.');
      }
    } else if (app.status === ST.CANCEL) {
      if (want === false) {
        applied = false;
      } else {
        sh.getRange(app.row, 3, 1, 6).setValues(
          [[new Date(), ST.APPLY, extra.car, extra.meal, extra.teams, extra.memo]]);
        log_(name, '신청', slotId);
        applied = true;
      }
    } else {
      // 신청 또는 대기
      if (want === true) {
        applied = true;
      } else {
        sh.getRange(app.row, 4).setValue(ST.CANCEL);
        log_(name, '취소', slotId);
        applied = false;
      }
    }

    SpreadsheetApp.flush();
    return { ok: true, applied: applied, count: countApplied_(loadApps_(), slotId) };
  });
}

function apiUpdateExtra_(req) {
  var name = auth_(req.token);
  var slotId = String(req.slotId || '');
  var slot = checkOpen_(slotId);
  var extra = cleanExtra_(req.extra, slot);

  return withLock_(function () {
    var app = findApp_(loadApps_(), slotId, name);
    if (!app || app.status === ST.CANCEL) {
      throw apiError_('NOT_APPLIED', '신청하지 않은 칸입니다. 화면을 새로 고쳐 주세요.');
    }
    sheet_(TAB.APPLY).getRange(app.row, 5, 1, 4).setValues(
      [[extra.car, extra.meal, extra.teams, extra.memo]]);
    log_(name, '정보수정', slotId);
    return { ok: true };
  });
}

// ==================== Status.gs ====================
/**
 * 신청현황, 추천결과, 전도인별 탭 그리기와 자동 갱신
 *
 * 배정 상태의 원본은 항상 신청 탭의 "상태" 열입니다.
 * 이 탭들은 원본을 읽어 보여 주기만 하므로 다시 그려도 배정이 사라지지 않습니다.
 */

/** 메뉴: 신청현황 새로고침 */
function refreshStatus() {
  refreshAll_(true);
}

/** 10분마다 실행되는 시간 기반 트리거. 바뀐 내용이 없으면 갱신 시각만 고칩니다. */
function autoRefresh() {
  try {
    refreshAll_(false);
  } catch (err) {
    if (!(err && err.apiCode === 'BUSY')) throw err;
  }
}

/** 메뉴: 자동 갱신 켜기 */
function enableAutoRefresh() {
  deleteTriggers_('autoRefresh');
  ScriptApp.newTrigger('autoRefresh').timeBased().everyMinutes(10).create();
  alert_('자동 갱신을 켰습니다. 10분마다 신청현황이 새로 그려집니다.');
}

/** 메뉴: 자동 갱신 끄기 */
function disableAutoRefresh() {
  var n = deleteTriggers_('autoRefresh');
  alert_(n ? '자동 갱신을 껐습니다.' : '켜져 있는 자동 갱신이 없습니다.');
}

function refreshAll_(force) {
  withLock_(function () {
    var data = loadAll_();
    var sig = computeSig_(data);
    var props = PropertiesService.getScriptProperties();
    var stamp = '마지막 갱신: ' + fmt_(new Date(), 'yyyy-MM-dd HH:mm');

    if (!force && props.getProperty('STATUS_SIG') === sig) {
      sheet_(TAB.STATUS).getRange(1, 1).setValue(stamp);
      return;
    }
    var view = buildView_(data, { title: stamp });
    renderView_(sheet_(TAB.STATUS), view, true);
    renderPersons_(data);
    props.setProperty('STATUS_SIG', sig);
  });
}

function computeSig_(data) {
  return sha256Hex_(JSON.stringify([data.slots, data.apps, data.roster.map(function (p) {
    return [p.name, p.gender];
  })]));
}

function headerColor_(applied, need, assigned) {
  if (need > 0 && assigned >= need) return COLOR.HEAD_FULL;
  if (need > 0 && applied < need) return COLOR.HEAD_SHORT;
  return COLOR.HEAD_NORMAL;
}

function blankRow_() {
  var row = [];
  for (var i = 0; i < VIEW.WIDTH; i++) row.push('');
  return row;
}

function fillRow_(value) {
  var row = [];
  for (var i = 0; i < VIEW.WIDTH; i++) row.push(value);
  return row;
}

/**
 * 시간대별 현황 표를 메모리에서 만듭니다.
 * opts.proposed: { '칸ID|이름': true } 자동 추천으로 새로 배정하려는 건 (추천결과 탭용)
 * opts.reco: true이면 체크박스 대신 "추천"/"기존" 글자를 씁니다.
 */
function buildView_(data, opts) {
  opts = opts || {};
  var proposed = opts.proposed || {};
  var gender = {};
  data.roster.forEach(function (p) { gender[p.name] = p.gender; });

  var bySlot = {}, totalApply = {}, totalAssign = {};
  data.apps.forEach(function (a) {
    if (a.status === ST.CANCEL) return;
    (bySlot[a.slotId] = bySlot[a.slotId] || []).push(a);
    totalApply[a.name] = (totalApply[a.name] || 0) + 1;
    if (a.status === ST.ASSIGN || proposed[appKey_(a.slotId, a.name)]) {
      totalAssign[a.name] = (totalAssign[a.name] || 0) + 1;
    }
  });

  var values = [], bgs = [], weights = [], checks = [];
  function push(row, bg, weight, isCheck) {
    values.push(row);
    bgs.push(fillRow_(bg));
    weights.push(fillRow_(weight));
    checks.push(isCheck);
  }

  var top = blankRow_();
  top[0] = opts.title || '';
  push(top, null, 'bold', false);

  var label = blankRow_();
  label[VIEW.CHECK - 1] = opts.reco ? '추천' : '배정';
  label[VIEW.NAME - 1] = '이름';
  label[VIEW.GENDER - 1] = '형제/자매';
  label[VIEW.CAR - 1] = '차량 제공';
  label[VIEW.MEAL - 1] = '식사 방식';
  label[VIEW.TEAMS - 1] = '대접 가능 팀 수';
  label[VIEW.MEMO - 1] = '메모';
  label[VIEW.TOTAL_APPLY - 1] = '총 신청 수';
  label[VIEW.TOTAL_ASSIGN - 1] = '총 배정 수';
  label[VIEW.TIME - 1] = '신청 시각';
  push(label, COLOR.LABEL, 'bold', false);

  data.slots.forEach(function (slot) {
    var list = (bySlot[slot.id] || []).slice().sort(function (a, b) {
      var ta = a.time instanceof Date ? a.time.getTime() : 0;
      var tb = b.time instanceof Date ? b.time.getTime() : 0;
      return ta - tb;
    });
    var assigned = 0;
    list.forEach(function (a) {
      if (a.status === ST.ASSIGN || proposed[appKey_(a.slotId, a.name)]) assigned++;
    });

    var head = blankRow_();
    head[0] = '■';
    head[1] = shortDate_(slot);
    head[2] = slot.item;
    head[3] = teamLabel_(slot.team);
    head[4] = slot.time;
    head[5] = slot.place;
    head[6] = '필요 ' + slot.need + '명';
    head[7] = '신청 ' + list.length + '명';
    head[8] = '배정 ' + assigned + '명';
    head[VIEW.KEY_SLOT - 1] = slot.id;
    head[VIEW.KEY_NAME - 1] = VIEW.HEADER_MARK;
    push(head, headerColor_(list.length, slot.need, assigned), 'bold', false);

    if (!list.length) {
      var none = blankRow_();
      none[VIEW.NAME - 1] = '(신청자 없음)';
      push(none, null, 'normal', false);
      return;
    }

    list.forEach(function (a) {
      var isProposed = !!proposed[appKey_(a.slotId, a.name)];
      var row = blankRow_();
      if (opts.reco) {
        row[VIEW.CHECK - 1] = isProposed ? '추천' : (a.status === ST.ASSIGN ? '기존' : '');
      } else {
        row[VIEW.CHECK - 1] = a.status === ST.ASSIGN;
      }
      row[VIEW.NAME - 1] = a.name;
      row[VIEW.GENDER - 1] = gender[a.name] || '';
      row[VIEW.CAR - 1] = a.car ? '차량' : '';
      row[VIEW.MEAL - 1] = a.meal;
      row[VIEW.TEAMS - 1] = a.teams === '' ? '' : a.teams + '팀';
      row[VIEW.MEMO - 1] = a.memo;
      row[VIEW.TOTAL_APPLY - 1] = String(totalApply[a.name] || 0);
      row[VIEW.TOTAL_ASSIGN - 1] = String(totalAssign[a.name] || 0);
      row[VIEW.TIME - 1] = a.time instanceof Date ? fmt_(a.time, 'MM/dd HH:mm') : '';
      row[VIEW.KEY_SLOT - 1] = a.slotId;
      row[VIEW.KEY_NAME - 1] = a.name;
      push(row, isProposed ? COLOR.RECO : null, 'normal', !opts.reco);
    });
  });

  return { values: values, bgs: bgs, weights: weights, checks: checks };
}

/** 만들어 둔 표를 시트에 그립니다. withCheckbox가 true이면 배정 열에 체크박스를 넣습니다. */
function renderView_(sh, view, withCheckbox) {
  var n = view.values.length;
  if (sh.getMaxColumns() < VIEW.WIDTH) {
    sh.insertColumnsAfter(sh.getMaxColumns(), VIEW.WIDTH - sh.getMaxColumns());
  }
  if (sh.getMaxRows() < n + 5) sh.insertRowsAfter(sh.getMaxRows(), n + 5 - sh.getMaxRows());

  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  // 날짜나 시간으로 자동 변환되지 않게 글자 서식으로 둡니다. (체크박스 열 제외)
  sh.getRange(1, 2, sh.getMaxRows(), VIEW.WIDTH - 1).setNumberFormat('@');

  var range = sh.getRange(1, 1, n, VIEW.WIDTH);
  range.setValues(view.values);
  range.setBackgrounds(view.bgs);
  range.setFontWeights(view.weights);
  range.setVerticalAlignment('middle');
  sh.getRange(1, 1, n, 1).setHorizontalAlignment('center');
  sh.getRange(1, VIEW.GENDER, n, 4).setHorizontalAlignment('center');
  sh.getRange(1, VIEW.TOTAL_APPLY, n, 3).setHorizontalAlignment('center');
  sh.getRange(1, 1).setHorizontalAlignment('left');

  var unprotected = [];
  if (withCheckbox) {
    var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    var rules = [];
    var start = 0;
    for (var i = 0; i <= n; i++) {
      var on = i < n && view.checks[i];
      if (i < n) rules.push([on ? rule : null]);
      if (on && !start) start = i + 1;
      if (!on && start) {
        unprotected.push(sh.getRange(start, 1, i + 1 - start, 1));
        start = 0;
      }
    }
    sh.getRange(1, 1, n, 1).setDataValidations(rules);
  }

  var widths = [50, 110, 80, 80, 80, 180, 220, 90, 90, 110, 120, 120];
  for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c + 1, widths[c]);
  sh.hideColumns(VIEW.KEY_SLOT, 2);
  sh.setFrozenRows(2);

  // 체크박스 열 외에는 고치려 하면 경고가 뜨게 합니다.
  try {
    var prots = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var prot = prots.length ? prots[0] : sh.protect();
    prot.setDescription('자동으로 만들어지는 탭입니다. 배정 체크박스만 고쳐 주세요.');
    prot.setWarningOnly(true);
    prot.setUnprotectedRanges(unprotected);
  } catch (err) {
    console.log('보호 범위 설정을 건너뜁니다: ' + err);
  }
}

/** 전도인별 요약 탭 */
function renderPersons_(data) {
  var slotById = {};
  data.slots.forEach(function (s) { slotById[s.id] = s; });
  var order = {};
  data.slots.forEach(function (s, i) { order[s.id] = i; });

  var people = {}, names = [];
  function person(name, gender) {
    if (!people[name]) {
      people[name] = { name: name, gender: gender || '', applied: [], assigned: [] };
      names.push(name);
    }
    return people[name];
  }
  data.roster.forEach(function (p) { person(p.name, p.gender); });

  data.apps.forEach(function (a) {
    if (a.status === ST.CANCEL) return;
    var p = person(a.name, '');
    p.applied.push(a.slotId);
    if (a.status === ST.ASSIGN) p.assigned.push(a.slotId);
  });

  function labels(ids) {
    return ids.slice().sort(function (x, y) {
      return (order[x] == null ? 9999 : order[x]) - (order[y] == null ? 9999 : order[y]);
    }).map(function (id) {
      return slotById[id] ? slotLabel_(slotById[id]) : id;
    }).join(', ');
  }

  var rows = names.map(function (name) {
    var p = people[name];
    return [p.name, p.gender, labels(p.applied), labels(p.assigned), p.applied.length, p.assigned.length];
  });

  var sh = sheet_(TAB.PERSON);
  if (sh.getMaxRows() > 1) {
    sh.getRange(2, 1, sh.getMaxRows() - 1, HEAD.PERSON.length).clearContent();
  }
  if (rows.length) {
    if (sh.getMaxRows() < rows.length + 1) {
      sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    }
    var range = sh.getRange(2, 1, rows.length, HEAD.PERSON.length);
    range.setValues(rows);
    range.setWrap(true).setVerticalAlignment('top');
  }
}

// ==================== Sync.gs ====================
/**
 * 배정 체크박스 동기화 (설치형 onEdit)와 "대기" 표시
 */

/** 메뉴: 배정 동기화 트리거 설치 */
function installSyncTrigger() {
  deleteTriggers_('onEditSync');
  ScriptApp.newTrigger('onEditSync').forSpreadsheet(ss_()).onEdit().create();
  alert_('배정 동기화를 켰습니다.\n이제 신청현황 탭에서 체크박스를 누르면 신청 탭의 상태가 함께 바뀝니다.');
}

/**
 * 신청현황 탭의 배정 체크박스가 바뀌면 신청 탭의 상태를 맞춥니다.
 *   체크함   → "배정"
 *   체크 해제 → "배정"이던 건만 "신청"으로 (대기, 취소는 그대로 둠)
 */
function onEditSync(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== TAB.STATUS) return;
  if (e.range.getColumn() !== VIEW.CHECK) return;
  var r1 = Math.max(e.range.getRow(), VIEW.FIRST_ROW);
  var r2 = e.range.getLastRow();
  if (r2 < r1) return;

  withLock_(function () {
    var lastRow = sh.getLastRow();
    if (lastRow < VIEW.FIRST_ROW) return;
    r2 = Math.min(r2, lastRow);
    var view = sh.getRange(1, 1, lastRow, VIEW.WIDTH).getValues();

    var data = loadAll_();
    var byKey = {};
    data.apps.forEach(function (a) { byKey[appKey_(a.slotId, a.name)] = a; });

    var applySheet = sheet_(TAB.APPLY);
    var touchedSlots = {}, logs = [], now = new Date();
    for (var r = r1; r <= r2; r++) {
      var row = view[r - 1];
      var slotId = String(row[VIEW.KEY_SLOT - 1]);
      var name = String(row[VIEW.KEY_NAME - 1]);
      if (!slotId || !name || name === VIEW.HEADER_MARK) continue;
      var app = byKey[appKey_(slotId, name)];
      if (!app || app.status === ST.CANCEL) continue;

      var checked = row[VIEW.CHECK - 1] === true;
      var next = null;
      if (checked && app.status !== ST.ASSIGN) next = ST.ASSIGN;
      if (!checked && app.status === ST.ASSIGN) next = ST.APPLY;
      if (!next) continue;

      applySheet.getRange(app.row, 4).setValue(next);
      app.status = next;
      touchedSlots[slotId] = true;
      logs.push([now, name, next === ST.ASSIGN ? '배정(장로)' : '배정해제(장로)', slotId]);
    }
    if (!logs.length) return;
    logMany_(logs);

    // 같은 사람의 "총 배정 수"와 블록 머리줄의 "현재 배정 수", 색을 바로 고칩니다.
    var totalAssign = {}, slotAssign = {}, slotApply = {}, needOf = {};
    data.slots.forEach(function (s) { needOf[s.id] = s.need; });
    data.apps.forEach(function (a) {
      if (a.status === ST.CANCEL) return;
      slotApply[a.slotId] = (slotApply[a.slotId] || 0) + 1;
      if (a.status === ST.ASSIGN) {
        totalAssign[a.name] = (totalAssign[a.name] || 0) + 1;
        slotAssign[a.slotId] = (slotAssign[a.slotId] || 0) + 1;
      }
    });

    var col = [];
    for (var i = VIEW.FIRST_ROW; i <= lastRow; i++) {
      var v = view[i - 1];
      var sid = String(v[VIEW.KEY_SLOT - 1]);
      var nm = String(v[VIEW.KEY_NAME - 1]);
      if (nm === VIEW.HEADER_MARK) {
        col.push(['배정 ' + (slotAssign[sid] || 0) + '명']);
        if (touchedSlots[sid]) {
          sh.getRange(i, 1, 1, VIEW.WIDTH).setBackground(
            headerColor_(slotApply[sid] || 0, needOf[sid] || 0, slotAssign[sid] || 0));
        }
      } else if (sid && nm) {
        col.push([String(totalAssign[nm] || 0)]);
      } else {
        col.push([v[VIEW.TOTAL_ASSIGN - 1]]);
      }
    }
    sh.getRange(VIEW.FIRST_ROW, VIEW.TOTAL_ASSIGN, col.length, 1).setValues(col);

    // 방금 바꾼 내용 때문에 자동 갱신이 표를 다시 그리지 않게 합니다.
    PropertiesService.getScriptProperties().setProperty('STATUS_SIG', computeSig_(data));
  });
}

/** 메뉴: 체크 안 된 신청을 모두 대기로 표시 */
function markWaiting() {
  if (!confirm_('상태가 "신청"인 건을 모두 "대기"로 바꿉니다.\n배정을 다 마친 뒤에 실행해 주세요. 계속할까요?')) return;

  var n = withLock_(function () {
    var sh = sheet_(TAB.APPLY);
    var apps = loadApps_();
    var count = 0;
    apps.forEach(function (a) {
      if (a.status === ST.APPLY) {
        sh.getRange(a.row, 4).setValue(ST.WAIT);
        count++;
      }
    });
    if (count) log_('(관리자)', '대기 표시 ' + count + '건', '');
    return count;
  });

  refreshAll_(true);
  alert_(n + '건을 "대기"로 바꿨습니다.\n\n' +
    '이제 설정 탭의 "결과공개여부"를 체크하면 전도인 화면에 배정 결과가 보입니다.');
}

// ==================== Recommend.gs ====================
/**
 * 자동 추천 배정
 *
 * 1. 이미 "배정"인 건은 그대로 둡니다.
 * 2. 신청자가 적은 칸부터 처리합니다.
 * 3. 지금까지 배정 수가 적은 사람을 우선합니다.
 * 4. 같으면 전체 신청 칸 수가 적은 사람, 그래도 같으면 무작위입니다.
 * 5. 필요 인원이 2명 이상인 봉사 칸에는 차량 제공자 1명, 자매 1명을 먼저 넣으려고 시도합니다.
 *    일정 탭에 "추천성별"을 적은 칸은 그 성별의 신청자만 추천합니다.
 * 6. 필요 인원이 차면 멈춥니다.
 * 추가: 같은 날 같은 항목의 같은 시간에 한 사람을 겹쳐 넣지 않습니다.
 */

/** 메뉴: 자동 추천 배정 (미리보기 생성). 신청 탭은 바꾸지 않습니다. */
function recommend() {
  var count = withLock_(function () {
    var data = loadAll_();
    var proposed = computeRecommendation_(data);
    var title = '자동 추천 미리보기 (' + fmt_(new Date(), 'yyyy-MM-dd HH:mm') +
      ' 생성) - 노란 줄이 새로 추천된 배정입니다. "추천 결과 반영" 메뉴를 눌러야 실제로 반영됩니다.';
    var view = buildView_(data, { title: title, proposed: proposed, reco: true });
    renderView_(sheet_(TAB.RECO), view, false);
    return Object.keys(proposed).length;
  });
  ss_().setActiveSheet(sheet_(TAB.RECO));
  alert_('추천 ' + count + '건을 "추천결과" 탭에 만들었습니다.\n' +
    '내용을 확인하신 뒤 "추천 결과 반영" 메뉴를 눌러 주세요.\n' +
    '마음에 들지 않으면 다시 "자동 추천 배정"을 누르면 새로 계산합니다.');
}

function computeRecommendation_(data) {
  var gender = {};
  data.roster.forEach(function (p) { gender[p.name] = p.gender; });
  var slotById = {};
  data.slots.forEach(function (s) { slotById[s.id] = s; });

  var bySlot = {}, totalApply = {}, assignCount = {}, busy = {};
  function slotTimeKey(slot) { return slot.date + '|' + slot.item + '|' + slot.time; }
  function markBusy(name, slot) { busy[name + '|' + slotTimeKey(slot)] = slot.id; }

  data.apps.forEach(function (a) {
    if (a.status === ST.CANCEL) return;
    (bySlot[a.slotId] = bySlot[a.slotId] || []).push(a);
    totalApply[a.name] = (totalApply[a.name] || 0) + 1;
    if (a.status === ST.ASSIGN) {
      assignCount[a.name] = (assignCount[a.name] || 0) + 1;
      if (slotById[a.slotId]) markBusy(a.name, slotById[a.slotId]);
    }
  });

  // 신청자가 적은 칸부터
  var slots = data.slots.filter(function (s) { return s.need > 0; }).map(function (s) {
    return { slot: s, n: (bySlot[s.id] || []).length, rnd: Math.random() };
  });
  slots.sort(function (a, b) { return a.n - b.n || a.rnd - b.rnd; });

  var proposed = {};
  slots.forEach(function (entry) {
    var slot = entry.slot;
    var list = bySlot[slot.id] || [];
    var fixed = list.filter(function (a) { return a.status === ST.ASSIGN; });
    var remaining = slot.need - fixed.length;
    if (remaining <= 0) return;

    var pool = list.filter(function (a) { return a.status !== ST.ASSIGN; }).map(function (a) {
      return { app: a, rnd: Math.random() };
    });

    function available(c) {
      if (slot.prefer && gender[c.app.name] !== slot.prefer) return false;
      var other = busy[c.app.name + '|' + slotTimeKey(slot)];
      return !other || other === slot.id;
    }
    function compare(x, y) {
      return (assignCount[x.app.name] || 0) - (assignCount[y.app.name] || 0) ||
        (totalApply[x.app.name] || 0) - (totalApply[y.app.name] || 0) ||
        x.rnd - y.rnd;
    }
    function pick(filter) {
      if (remaining <= 0) return false;
      var cands = pool.filter(function (c) { return available(c) && (!filter || filter(c)); });
      if (!cands.length) return false;
      cands.sort(compare);
      var chosen = cands[0];
      pool.splice(pool.indexOf(chosen), 1);
      proposed[appKey_(slot.id, chosen.app.name)] = true;
      assignCount[chosen.app.name] = (assignCount[chosen.app.name] || 0) + 1;
      markBusy(chosen.app.name, slot);
      picked.push(chosen.app);
      remaining--;
      return true;
    }

    var picked = [];
    if (isService_(slot) && slot.need >= 2 && !slot.prefer) {
      var hasCar = fixed.some(function (a) { return a.car; });
      if (!hasCar) pick(function (c) { return c.app.car; });
      var hasSister = fixed.concat(picked).some(function (a) { return gender[a.name] === '자매'; });
      if (!hasSister) pick(function (c) { return gender[c.app.name] === '자매'; });
    }
    while (pick(null)) { /* 필요 인원이 찰 때까지 */ }
  });

  return proposed;
}

/** 메뉴: 추천 결과 반영. 추천결과 탭의 "추천" 줄을 신청 탭에 "배정"으로 적습니다. */
function applyRecommendation() {
  var sh = sheet_(TAB.RECO);
  var lastRow = sh.getLastRow();
  var keys = [];
  if (lastRow >= VIEW.FIRST_ROW) {
    var view = sh.getRange(1, 1, lastRow, VIEW.WIDTH).getValues();
    for (var r = VIEW.FIRST_ROW; r <= lastRow; r++) {
      var row = view[r - 1];
      if (row[VIEW.CHECK - 1] === '추천' && row[VIEW.KEY_SLOT - 1] && row[VIEW.KEY_NAME - 1]) {
        keys.push(appKey_(String(row[VIEW.KEY_SLOT - 1]), String(row[VIEW.KEY_NAME - 1])));
      }
    }
  }
  if (!keys.length) {
    alert_('반영할 추천이 없습니다. 먼저 "자동 추천 배정 (미리보기 생성)"을 실행해 주세요.');
    return;
  }
  if (!confirm_('추천 ' + keys.length + '건을 "배정"으로 반영합니다. 계속할까요?')) return;

  var result = withLock_(function () {
    var applySheet = sheet_(TAB.APPLY);
    var byKey = {};
    loadApps_().forEach(function (a) { byKey[appKey_(a.slotId, a.name)] = a; });
    var done = 0, skipped = 0, logs = [], now = new Date();
    keys.forEach(function (key) {
      var app = byKey[key];
      // 미리보기를 만든 뒤 취소했거나 이미 배정된 건은 건너뜁니다.
      if (!app || (app.status !== ST.APPLY && app.status !== ST.WAIT)) { skipped++; return; }
      applySheet.getRange(app.row, 4).setValue(ST.ASSIGN);
      logs.push([now, app.name, '배정(자동추천)', app.slotId]);
      done++;
    });
    logMany_(logs);
    return { done: done, skipped: skipped };
  });

  // 같은 추천이 두 번 반영되지 않게 미리보기를 비웁니다.
  sh.clear();
  sh.getRange(1, 1).setValue('추천 결과를 반영했습니다. (' + fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ')');
  refreshAll_(true);
  ss_().setActiveSheet(sheet_(TAB.STATUS));
  alert_(result.done + '건을 배정으로 반영했습니다.' +
    (result.skipped ? '\n그 사이 취소되었거나 이미 배정되어 건너뛴 건: ' + result.skipped + '건' : '') +
    '\n신청현황 탭에서 확인하고 필요하면 체크박스로 고쳐 주세요.');
}

// ==================== TestData.gs ====================
/**
 * 테스트 데이터 넣기 / 지우기
 *
 * 가짜 이름에는 "(테스트)"가, 가짜 칸ID에는 "TEST-"가 붙습니다.
 * 지울 때는 이 표시가 있는 행만 지우므로 실제 자료는 건드리지 않습니다.
 *
 * 테스트 로그인: 이름 "김하늘(테스트)", 전화번호 010-0000-0001
 *               (두 번째 사람은 010-0000-0002, 이런 식으로 이어집니다)
 */

var TEST_MARK = '(테스트)';
var TEST_PREFIX = 'TEST-';

/** 메뉴: 테스트 데이터 넣기 */
function insertTestData() {
  withLock_(function () {
    removeTestRows_();

    var people = [
      ['김하늘', '형제'], ['이바다', '자매'], ['박나무', '형제'], ['최구름', '자매'],
      ['정달빛', '자매'], ['강산들', '형제'], ['조새벽', '자매'], ['윤이슬', '자매'],
      ['장노을', '형제'], ['임보람', '자매'], ['한솔잎', '형제'], ['오별님', '자매']
    ];
    var roster = people.map(function (p, i) {
      var phone = '0100000' + ('0000' + (i + 1)).slice(-4);
      return [p[0] + TEST_MARK, hashPhone_(phone), p[1]];
    });
    appendRows_(sheet_(TAB.ROSTER), roster);

    var days = [['2026-12-29', '1229'], ['2026-12-30', '1230'], ['2026-12-31', '1231'], ['2027-01-01', '0101']];
    var slots = [];
    days.forEach(function (d) {
      slots.push([TEST_PREFIX + d[1] + '-AM-A', d[0], '', '오전봉사', 'A', '09:30', '왕국회관', 3, true, '']);
      slots.push([TEST_PREFIX + d[1] + '-AM-B', d[0], '', '오전봉사', 'B', '09:30', '1구역 모임 장소', 3, true, '']);
      slots.push([TEST_PREFIX + d[1] + '-PM-A', d[0], '', '오후봉사', 'A', '14:00', '왕국회관', 3, true, '']);
      slots.push([TEST_PREFIX + d[1] + '-PM-B', d[0], '', '오후봉사', 'B', '14:00', '2구역 모임 장소', 3, true, '']);
      slots.push([TEST_PREFIX + d[1] + '-DN', d[0], '', '저녁식사', '공통', '18:00', '', 1, true, '']);
    });
    appendRows_(sheet_(TAB.SCHEDULE), slots);

    var apps = [], base = new Date().getTime() - 3600000;
    slots.forEach(function (s) {
      var dinner = s[3] === '저녁식사';
      roster.forEach(function (p) {
        if (Math.random() > (dinner ? 0.25 : 0.4)) return;
        base += 1000 * Math.floor(1 + Math.random() * 120);
        apps.push([
          s[0], p[0], new Date(base), ST.APPLY,
          !dinner && Math.random() < 0.3 ? 'O' : '',
          dinner ? (Math.random() < 0.6 ? '집' : '외식') : '',
          dinner ? (Math.random() < 0.5 ? 1 : 2) : '',
          Math.random() < 0.1 ? '테스트 메모입니다' : ''
        ]);
      });
    });
    var applySheet = sheet_(TAB.APPLY);
    if (apps.length) {
      applySheet.getRange(applySheet.getLastRow() + 1, 1, apps.length, HEAD.APPLY.length).setValues(apps);
    }

    var conf = sheet_(TAB.CONFIG);
    if (!(conf.getRange(2, 2).getValue() instanceof Date)) {
      conf.getRange(2, 2).setValue(new Date(new Date().getTime() + 7 * 24 * 3600000));
    }
  });

  refreshAll_(true);
  alert_('테스트 데이터를 넣었습니다.\n\n' +
    '웹 화면 테스트 로그인\n  이름: 김하늘(테스트)\n  전화번호: 010-0000-0001\n\n' +
    '실제로 쓰기 전에 꼭 "테스트 데이터 지우기"를 실행해 주세요.');
}

/** 메뉴: 테스트 데이터 지우기 */
function removeTestData() {
  if (!confirm_('"(테스트)" 이름과 "TEST-" 칸ID가 붙은 행을 명단, 일정, 신청, 로그 탭에서 지웁니다. 계속할까요?')) return;
  var n = withLock_(removeTestRows_);
  refreshAll_(true);
  alert_('테스트 데이터 ' + n + '줄을 지웠습니다.');
}

function removeTestRows_() {
  function isTestName(v) { return String(v).indexOf(TEST_MARK) >= 0; }
  function isTestSlot(v) { return String(v).indexOf(TEST_PREFIX) === 0; }
  var n = 0;
  n += filterRows_(sheet_(TAB.ROSTER), HEAD.ROSTER.length, function (r) { return isTestName(r[0]); });
  n += filterRows_(sheet_(TAB.SCHEDULE), HEAD.SCHEDULE.length, function (r) { return isTestSlot(r[0]); });
  n += filterRows_(sheet_(TAB.APPLY), HEAD.APPLY.length, function (r) { return isTestSlot(r[0]) || isTestName(r[1]); });
  n += filterRows_(sheet_(TAB.LOG), HEAD.LOG.length, function (r) { return isTestName(r[1]) || isTestSlot(r[3]); });
  return n;
}

/** 조건에 맞는 행을 빼고 나머지를 위로 당겨 다시 씁니다. 지운 줄 수를 돌려줍니다. */
function filterRows_(sh, width, isRemoved) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var range = sh.getRange(2, 1, last - 1, width);
  var values = range.getValues();
  var keep = values.filter(function (r) { return !isRemoved(r); });
  var removed = values.length - keep.length;
  if (!removed) return 0;
  range.clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, width).setValues(keep);
  return removed;
}
