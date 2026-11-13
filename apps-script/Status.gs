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
    head[3] = slot.team === '공통' ? '공통' : slot.team + '팀';
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
