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
