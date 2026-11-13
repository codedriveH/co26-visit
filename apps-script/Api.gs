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
