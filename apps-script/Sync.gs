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
