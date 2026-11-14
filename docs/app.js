/* 순회 방문 신청 - 화면 동작 */
(function () {
  'use strict';

  var API_URL = (window.APP_CONFIG && window.APP_CONFIG.API_URL) || '';
  var TOKEN_KEY = 'co_visit_token';
  var ITEM_ORDER = ['오전봉사', '점심식사', '오후봉사', '저녁식사'];
  var ITEM_NAME = { '오전봉사': '오전 봉사', '점심식사': '점심 식사', '오후봉사': '오후 봉사', '저녁식사': '저녁 식사' };

  var state = {
    token: '',
    data: null,      // 서버에서 받은 자료
    pending: {},     // 서버 응답을 기다리는 칸
    loadedAt: 0
  };

  function $(id) { return document.getElementById(id); }

  /* ---------- 저장소 (막혀 있어도 화면은 동작해야 함) ---------- */

  function loadToken() {
    try { return window.localStorage.getItem(TOKEN_KEY) || ''; } catch (e) { return ''; }
  }
  function saveToken(token) {
    try {
      if (token) window.localStorage.setItem(TOKEN_KEY, token);
      else window.localStorage.removeItem(TOKEN_KEY);
    } catch (e) { /* 저장이 안 되면 다음에 다시 로그인하면 됩니다 */ }
  }

  /* ---------- 서버 요청 ---------- */

  function api(action, payload) {
    var body = { action: action };
    var key;
    for (key in payload) {
      if (Object.prototype.hasOwnProperty.call(payload, key)) body[key] = payload[key];
    }
    if (state.token && action !== 'login') body.token = state.token;

    var controller = window.AbortController ? new AbortController() : null;
    var timer = setTimeout(function () { if (controller) controller.abort(); }, 30000);

    // 사전 요청(preflight)을 피하려고 text/plain으로 보냅니다.
    return fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body),
      redirect: 'follow',
      signal: controller ? controller.signal : undefined
    }).then(function (res) {
      return res.json();
    }).then(function (json) {
      clearTimeout(timer);
      if (!json || !json.ok) {
        var err = new Error((json && json.message) || '문제가 생겼습니다. 잠시 후 다시 시도해 주세요.');
        err.code = (json && json.code) || 'ERROR';
        throw err;
      }
      return json;
    }, function () {
      clearTimeout(timer);
      var err = new Error('연결이 잘 되지 않습니다. 인터넷을 확인하신 뒤 다시 시도해 주세요.');
      err.code = 'NETWORK';
      throw err;
    });
  }

  /* ---------- 공통 화면 요소 ---------- */

  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
  }

  function showView(name) {
    $('loginView').hidden = name !== 'login';
    $('mainView').hidden = name !== 'main';
  }

  function showLoading(text) {
    $('loadingText').textContent = text || '불러오는 중입니다';
    $('loading').hidden = false;
  }
  function hideLoading() { $('loading').hidden = true; }

  var toastTimer = null;
  function toast(message) {
    var box = $('toast');
    box.textContent = message;
    box.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { box.hidden = true; }, 4500);
  }

  /* ---------- 로그인 ---------- */

  function showLogin(message) {
    state.token = '';
    state.data = null;
    saveToken('');
    closeSheet();
    showView('login');
    var err = $('loginError');
    err.hidden = !message;
    err.textContent = message || '';
  }

  function onLoginSubmit(ev) {
    ev.preventDefault();
    var name = $('nameInput').value.replace(/\s+/g, '');
    var phone = $('phoneInput').value.replace(/\D/g, '');
    var err = $('loginError');

    if (!API_URL) {
      err.textContent = '아직 준비 중입니다. 잠시 후 다시 접속해 주세요.';
      err.hidden = false;
      return;
    }
    if (!name) { err.textContent = '이름을 입력해 주세요.'; err.hidden = false; $('nameInput').focus(); return; }
    if (phone.length < 10) {
      err.textContent = '휴대폰 번호를 끝까지 입력해 주세요.';
      err.hidden = false;
      $('phoneInput').focus();
      return;
    }
    err.hidden = true;
    document.activeElement && document.activeElement.blur();

    showLoading('확인하는 중입니다');
    api('login', { name: name, phone: phone }).then(function (res) {
      state.token = res.token;
      saveToken(res.token);
      $('phoneInput').value = '';
      return loadData(true);
    }).catch(function (e) {
      hideLoading();
      err.textContent = e.message;
      err.hidden = false;
    });
  }

  function logout() {
    showLogin('');
  }

  /* ---------- 자료 불러오기 ---------- */

  function loadData(withOverlay) {
    if (withOverlay) showLoading('불러오는 중입니다');
    return api('getData', {}).then(function (res) {
      state.data = res;
      state.loadedAt = Date.now();
      hideLoading();
      showView('main');
      render();
    }).catch(function (e) {
      hideLoading();
      if (e.code === 'AUTH') {
        showLogin(state.data ? e.message : '');
        return;
      }
      if (state.data) {
        toast(e.message);
      } else {
        showView('main');
        renderError(e.message);
      }
    });
  }

  function renderError(message) {
    $('userName').textContent = '';
    $('myCount').textContent = '';
    $('deadline').textContent = '';
    $('resultBox').hidden = true;
    $('legend').hidden = true;
    $('days').textContent = '';
    var notice = $('notice');
    notice.hidden = false;
    notice.textContent = message + ' 위의 "새로 고침"을 눌러 주세요.';
  }

  /* ---------- 신청 화면 그리기 ---------- */

  function isReadOnly() {
    return !!(state.data && (state.data.closed || state.data.published));
  }
  function isService(slot) { return slot.item.indexOf('봉사') >= 0; }
  function isDinner(slot) { return slot.item.indexOf('식사') >= 0; }

  function dateText(slot) {
    var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(slot.date);
    var base = m ? Number(m[2]) + '월 ' + Number(m[3]) + '일' : slot.date;
    return base + (slot.dow ? ' (' + slot.dow + ')' : '');
  }
  function itemText(slot) { return ITEM_NAME[slot.item] || slot.item; }
  // A, B는 "A팀"으로, 그 밖의 값(함께할 사람 이름 등)은 그대로 보여 줍니다.
  function teamText(slot) {
    if (!slot.team || slot.team === '공통') return '';
    return /^[A-Za-z]$/.test(slot.team) ? slot.team + '팀' : slot.team;
  }
  function slotTitle(slot) {
    return dateText(slot) + ' ' + itemText(slot) + (teamText(slot) ? ' ' + teamText(slot) : '');
  }
  function slotSub(slot) {
    var parts = [];
    if (slot.time) parts.push(slot.time);
    if (slot.place) parts.push(slot.place);
    return parts.join(' · ');
  }

  function extraText(slot, mine) {
    var parts = [];
    if (isService(slot) && mine.car) parts.push('차량 제공');
    if (isDinner(slot)) {
      if (mine.meal) parts.push(mine.meal === '집' ? '집에서' : '외식');
    }
    return parts.join(' · ');
  }

  function stateInfo(slot, mine) {
    if (mine) {
      if (mine.status === '배정') return { cls: 'is-assigned', mark: '★', text: '배정됨' };
      if (mine.status === '대기') return { cls: 'is-wait', mark: '…', text: '대기' };
      return { cls: 'is-applied', mark: '✓', text: '신청됨' };
    }
    if (isReadOnly()) return { cls: 'is-off', mark: '', text: '신청 안 함' };
    if (!slot.open) return { cls: 'is-off', mark: '', text: '신청 안 받음' };
    return { cls: 'is-empty', mark: '＋', text: '신청하기' };
  }

  function render() {
    var data = state.data;
    if (!data) return;
    var readOnly = isReadOnly();

    var mineCount = 0, id;
    for (id in data.mine) {
      if (Object.prototype.hasOwnProperty.call(data.mine, id)) mineCount++;
    }
    $('userName').textContent = data.name + ' 님';
    $('myCount').textContent = '내 신청 ' + mineCount + '건';
    var dl = $('deadline');
    if (data.closed) dl.textContent = '신청이 마감되었습니다';
    else if (data.deadlineText) dl.textContent = '신청 마감: ' + data.deadlineText;
    else dl.textContent = '';
    dl.className = 'topbar-deadline' + (data.closed ? ' is-closed' : '');

    var notice = $('notice');
    if (data.published) {
      notice.hidden = true;
    } else if (data.closed) {
      notice.hidden = false;
      notice.textContent = '신청이 마감되었습니다. 배정 결과가 나오면 이 화면에서 보실 수 있습니다.';
    } else {
      notice.hidden = false;
      notice.textContent = '참여하실 수 있는 칸을 모두 눌러 신청해 주세요. 신청하신 내용을 보고 장로들이 배정합니다.';
    }

    renderResult(data);
    renderLegend(data);

    var days = $('days');
    days.textContent = '';
    if (!data.slots.length) {
      days.appendChild(el('p', 'empty', '아직 일정이 올라오지 않았습니다.'));
      return;
    }

    var byDate = {}, dateKeys = [];
    data.slots.forEach(function (slot) {
      if (!byDate[slot.date]) { byDate[slot.date] = []; dateKeys.push(slot.date); }
      byDate[slot.date].push(slot);
    });
    dateKeys.sort();

    dateKeys.forEach(function (date, dayIndex) {
      var list = byDate[date];
      // 날짜마다 다른 색으로 구분합니다. (6가지 색을 차례로 씁니다)
      var card = el('section', 'day day-c' + (dayIndex % 6));
      card.appendChild(el('h2', 'day-title', dateText(list[0])));

      var items = [];
      list.forEach(function (slot) { if (items.indexOf(slot.item) < 0) items.push(slot.item); });
      items.sort(function (a, b) {
        var ia = ITEM_ORDER.indexOf(a), ib = ITEM_ORDER.indexOf(b);
        return (ia < 0 ? 9 : ia) - (ib < 0 ? 9 : ib);
      });

      items.forEach(function (item) {
        var slots = list.filter(function (s) { return s.item === item; });
        var block = el('div', 'item');
        block.appendChild(el('h3', 'item-title', ITEM_NAME[item] || item));

        // 서버가 보낸 순서(시간, 일정 탭 순서)를 지키면서 같은 시간끼리 묶습니다.
        var times = [], byTime = {};
        slots.forEach(function (s) {
          if (!byTime[s.time]) { byTime[s.time] = []; times.push(s.time); }
          byTime[s.time].push(s);
        });
        times.forEach(function (time) {
          var group = byTime[time];
          var showTime = times.length > 1;
          if (showTime && time) block.appendChild(el('div', 'time-title', time));
          var grid = el('div', 'slots' + (group.length > 1 ? ' slots-2' : ''));
          group.forEach(function (slot) { grid.appendChild(renderSlot(slot, readOnly, showTime)); });
          block.appendChild(grid);
        });
        card.appendChild(block);
      });
      days.appendChild(card);
    });
  }

  function renderSlot(slot, readOnly, hideTime) {
    var mine = state.data.mine[slot.id];
    var info = stateInfo(slot, mine);
    var disabled = readOnly || (!mine && !slot.open);
    var btn = el('button', 'slot ' + info.cls + (state.pending[slot.id] ? ' is-pending' : ''));
    btn.type = 'button';
    btn.setAttribute('data-id', slot.id);
    if (disabled) {
      btn.disabled = true;
    }

    var head = el('span', 'slot-head');
    if (teamText(slot)) head.appendChild(el('span', 'slot-team', teamText(slot)));
    var sub = hideTime ? slot.place : slotSub(slot);
    if (sub) head.appendChild(el('span', 'slot-info', sub));
    btn.appendChild(head);

    var st = el('span', 'slot-state');
    if (state.pending[slot.id]) {
      st.appendChild(el('span', 'mini-spinner'));
    } else if (info.mark) {
      var mark = el('span', 'slot-mark', info.mark);
      mark.setAttribute('aria-hidden', 'true');
      st.appendChild(mark);
    }
    st.appendChild(el('span', '', info.text));
    btn.appendChild(st);

    if (mine && extraText(slot, mine)) btn.appendChild(el('span', 'slot-extra', extraText(slot, mine)));
    btn.appendChild(el('span', 'slot-count', '신청 ' + slot.count + '명'));
    return btn;
  }

  function renderLegend(data) {
    var box = $('legend');
    box.textContent = '';
    box.hidden = false;
    var items = data.published
      ? [['is-assigned', '★ 배정됨'], ['is-wait', '… 대기']]
      : [['is-empty', '＋ 신청 전'], ['is-applied', '✓ 신청됨']];
    items.forEach(function (it) { box.appendChild(el('span', 'legend-item ' + it[0], it[1])); });
  }

  function renderResult(data) {
    var box = $('resultBox');
    box.textContent = '';
    box.hidden = !data.published;
    if (!data.published) return;

    box.appendChild(el('h2', 'result-title', '내 배정'));
    var assigned = [], waiting = [];
    data.slots.forEach(function (slot) {
      var mine = data.mine[slot.id];
      if (!mine) return;
      (mine.status === '배정' ? assigned : waiting).push(slot);
    });

    if (!assigned.length && !waiting.length) {
      box.appendChild(el('p', 'result-empty', '신청하신 내용이 없습니다.'));
      return;
    }
    if (!assigned.length) {
      box.appendChild(el('p', 'result-empty', '이번에는 배정된 칸이 없습니다.'));
    }
    function addList(list, cls, mark, label) {
      if (!list.length) return;
      var ul = el('ul', 'result-list');
      list.forEach(function (slot) {
        var li = el('li', 'result-item ' + cls);
        li.appendChild(el('span', 'result-badge', mark + ' ' + label));
        var text = el('span', 'result-text');
        text.appendChild(el('strong', '', slotTitle(slot)));
        if (slotSub(slot)) text.appendChild(el('span', 'result-sub', slotSub(slot)));
        li.appendChild(text);
        ul.appendChild(li);
      });
      box.appendChild(ul);
    }
    addList(assigned, 'is-assigned', '★', '배정');
    addList(waiting, 'is-wait', '…', '대기');
    box.appendChild(el('p', 'result-help', '"대기"는 이번에는 배정되지 않았다는 뜻입니다. 궁금하신 점은 장로에게 문의해 주세요.'));
  }

  /* ---------- 추가 정보 창 ---------- */

  var sheetState = null;

  function findSlot(id) {
    var slots = state.data ? state.data.slots : [];
    for (var i = 0; i < slots.length; i++) if (slots[i].id === id) return slots[i];
    return null;
  }

  function choiceGroup(label, options, value, onChange, hint) {
    var wrap = el('div', 'choice');
    wrap.appendChild(el('div', 'choice-label', label));
    if (hint) wrap.appendChild(el('div', 'choice-hint', hint));
    var row = el('div', 'choice-row');
    var buttons = [];
    options.forEach(function (opt) {
      var b = el('button', 'choice-btn', opt.text);
      b.type = 'button';
      b.setAttribute('aria-pressed', String(opt.value === value));
      b.addEventListener('click', function () {
        buttons.forEach(function (x) { x.setAttribute('aria-pressed', 'false'); });
        b.setAttribute('aria-pressed', 'true');
        onChange(opt.value);
      });
      buttons.push(b);
      row.appendChild(b);
    });
    wrap.appendChild(row);
    return wrap;
  }

  function openSheet(slot) {
    var mine = state.data.mine[slot.id];
    var extra = {
      car: mine ? !!mine.car : false,
      meal: mine && mine.meal ? mine.meal : '외식',
      teams: 2,
      memo: mine ? (mine.memo || '') : ''
    };
    sheetState = { slot: slot, extra: extra };

    $('sheetTitle').textContent = slotTitle(slot);
    $('sheetSub').textContent = slotSub(slot);
    var body = $('sheetBody');
    body.textContent = '';

    if (isService(slot)) {
      body.appendChild(choiceGroup('차량을 제공하실 수 있나요?',
        [{ text: '예, 가능합니다', value: true }, { text: '아니요', value: false }],
        extra.car, function (v) { extra.car = v; }));
    }
    if (isDinner(slot)) {
      body.appendChild(choiceGroup('식사는 어디에서 하시나요?',
        [{ text: '집에서', value: '집' }, { text: '외식', value: '외식' }],
        extra.meal, function (v) { extra.meal = v; }));
      // 식사는 항상 두 팀을 함께 대접하므로 팀 수는 묻지 않습니다.
    }

    var memoWrap = el('label', 'field');
    memoWrap.appendChild(el('span', 'field-label', '전하실 말씀 (쓰지 않으셔도 됩니다)'));
    var memo = document.createElement('input');
    memo.type = 'text';
    memo.maxLength = 200;
    memo.value = extra.memo;
    memo.addEventListener('input', function () { extra.memo = memo.value; });
    memoWrap.appendChild(memo);
    body.appendChild(memoWrap);

    var actions = $('sheetActions');
    actions.textContent = '';
    function addButton(text, cls, handler) {
      var b = el('button', 'btn ' + cls, text);
      b.type = 'button';
      b.addEventListener('click', handler);
      actions.appendChild(b);
    }
    if (mine) {
      addButton('바꾼 내용 저장', 'btn-primary', function () { closeSheet(); doUpdateExtra(slot, extra); });
      addButton('신청 취소하기', 'btn-danger', function () { closeSheet(); doToggle(slot, false, extra); });
    } else {
      addButton('신청하기', 'btn-primary', function () { closeSheet(); doToggle(slot, true, extra); });
    }
    addButton('닫기', 'btn-plain', closeSheet);

    $('sheetBackdrop').hidden = false;
    document.body.classList.add('no-scroll');
  }

  function closeSheet() {
    sheetState = null;
    $('sheetBackdrop').hidden = true;
    document.body.classList.remove('no-scroll');
  }

  /* ---------- 신청 / 취소 / 수정 (낙관적 업데이트) ---------- */

  function cloneMine(m) {
    return m ? { status: m.status, car: m.car, meal: m.meal, teams: m.teams, memo: m.memo } : null;
  }

  function mineFromExtra(slot, extra, status) {
    return {
      status: status,
      car: isService(slot) ? !!extra.car : false,
      meal: isDinner(slot) ? extra.meal : '',
      teams: isDinner(slot) && slot.team === '공통' ? extra.teams : '',
      memo: extra.memo || ''
    };
  }

  function handleFailure(e) {
    if (e.code === 'AUTH') {
      showLogin(e.message);
      return;
    }
    render();
    toast(e.message);
    if (e.code === 'CLOSED') loadData(false);
  }

  function doToggle(slot, on, extra) {
    if (state.pending[slot.id]) return;
    var data = state.data;
    var before = cloneMine(data.mine[slot.id]);
    var beforeCount = slot.count;

    // 먼저 화면을 바꾸고 서버에 보냅니다.
    if (on) {
      data.mine[slot.id] = mineFromExtra(slot, extra, '신청');
      if (!before) slot.count = beforeCount + 1;
    } else {
      delete data.mine[slot.id];
      if (before) slot.count = Math.max(0, beforeCount - 1);
    }
    state.pending[slot.id] = true;
    render();
    // 저장이 끝날 때까지 다른 버튼을 누르지 못하게 화면을 가립니다.
    showLoading(on ? '신청하는 중입니다' : '취소하는 중입니다');

    api('toggle', { slotId: slot.id, want: on ? 'on' : 'off', extra: extra }).then(function (res) {
      delete state.pending[slot.id];
      hideLoading();
      if (state.data !== data) return;
      slot.count = res.count;
      if (res.applied && !data.mine[slot.id]) data.mine[slot.id] = mineFromExtra(slot, extra, '신청');
      if (!res.applied) delete data.mine[slot.id];
      render();
      toast(res.applied ? '신청되었습니다.' : '신청을 취소했습니다.');
    }).catch(function (e) {
      delete state.pending[slot.id];
      hideLoading();
      if (state.data === data) {
        // 실패하면 원래대로 되돌립니다.
        if (before) data.mine[slot.id] = before; else delete data.mine[slot.id];
        slot.count = beforeCount;
      }
      handleFailure(e);
    });
  }

  function doUpdateExtra(slot, extra) {
    if (state.pending[slot.id]) return;
    var data = state.data;
    var before = cloneMine(data.mine[slot.id]);
    if (!before) return;

    data.mine[slot.id] = mineFromExtra(slot, extra, before.status);
    state.pending[slot.id] = true;
    render();
    showLoading('저장하는 중입니다');

    api('updateExtra', { slotId: slot.id, extra: extra }).then(function () {
      delete state.pending[slot.id];
      hideLoading();
      render();
      toast('저장되었습니다.');
    }).catch(function (e) {
      delete state.pending[slot.id];
      hideLoading();
      if (state.data === data) data.mine[slot.id] = before;
      handleFailure(e);
    });
  }

  function onSlotClick(ev) {
    var node = ev.target;
    while (node && node !== ev.currentTarget && !(node.getAttribute && node.getAttribute('data-id'))) {
      node = node.parentNode;
    }
    if (!node || node === ev.currentTarget) return;
    var slot = findSlot(node.getAttribute('data-id'));
    if (!slot || isReadOnly() || state.pending[slot.id]) return;

    var mine = state.data.mine[slot.id];
    if (!mine && !slot.open) return;
    if (!mine && !isService(slot) && !isDinner(slot)) {
      // 추가로 여쭐 것이 없는 칸은 바로 신청합니다.
      doToggle(slot, true, { memo: '' });
      return;
    }
    openSheet(slot);
  }

  /* ---------- 시작 ---------- */

  function init() {
    $('loginForm').addEventListener('submit', onLoginSubmit);
    $('logoutBtn').addEventListener('click', logout);
    $('refreshBtn').addEventListener('click', function () {
      if (state.token) loadData(true); else showLogin('');
    });
    $('days').addEventListener('click', onSlotClick);
    $('sheetBackdrop').addEventListener('click', function (ev) {
      if (ev.target === $('sheetBackdrop')) closeSheet();
    });

    // 다른 앱에 다녀온 뒤에는 최신 내용으로 다시 불러옵니다.
    document.addEventListener('visibilitychange', function () {
      var busy = false, id;
      for (id in state.pending) busy = true;
      if (document.visibilityState === 'visible' && state.token && state.data && !sheetState && !busy &&
          Date.now() - state.loadedAt > 60000) {
        loadData(false);
      }
    });

    state.token = loadToken();
    if (state.token && API_URL) {
      loadData(true);
    } else {
      showLogin('');
      if (!API_URL) {
        $('loginError').textContent = '아직 준비 중입니다. 잠시 후 다시 접속해 주세요.';
        $('loginError').hidden = false;
      }
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
