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
