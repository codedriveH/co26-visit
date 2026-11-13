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
  var sched = ensureSheet_(TAB.SCHEDULE, HEAD.SCHEDULE, [130, 110, 50, 90, 60, 80, 180, 80, 80]);
  sched.getRange(2, 1, rows, 1).setNumberFormat('@');
  sched.getRange(2, 2, rows, 1).setNumberFormat('yyyy-mm-dd');
  sched.getRange(2, 4, rows, 1).setDataValidation(
    dv().requireValueInList(ITEMS, true).setAllowInvalid(true).build());
  sched.getRange(2, 5, rows, 1).setDataValidation(
    dv().requireValueInList(TEAMS, true).setAllowInvalid(false).build());
  sched.getRange(2, 6, rows, 1).setNumberFormat('@');
  sched.getRange(2, 9, rows, 1).setDataValidation(dv().requireCheckbox().build());
  sched.getRange(1, 3).setNote('비워 두면 날짜에서 자동으로 계산합니다.');
  sched.getRange(1, 9).setNote('체크한 칸만 전도인이 신청할 수 있습니다.');

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
