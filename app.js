/* B4 레슨 노트 — 앱 화면 (학생 · 선생님)
 * 주소 규칙
 *   학생:   index.html?k=학생키
 *   선생님: index.html?admin=관리자키
 *   미리보기: config.js 의 API_URL 이 비어 있거나 ?demo=1
 */
(function () {
  'use strict';

  var CFG = window.B4_CONFIG || {};
  var API = String(CFG.API_URL || '').trim();
  var params = new URLSearchParams(location.search);
  var DEMO = !API || params.get('demo') === '1';
  var $app = document.getElementById('app');
  var $toast = document.getElementById('toast');

  /* ---------- 공통 유틸 ---------- */
  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  var store = {
    get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
    del: function (k) { try { localStorage.removeItem(k); } catch (e) {} }
  };
  var toastTimer = null;
  function toast(msg) {
    $toast.textContent = msg;
    $toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { $toast.hidden = true; }, 2600);
  }
  function fmtTime(s) {
    if (!isFinite(s) || s < 0) s = 0;
    var m = Math.floor(s / 60), r = Math.floor(s % 60);
    return m + ':' + (r < 10 ? '0' : '') + r;
  }
  function safeUrl(u) { return /^(https?:\/\/|\?)/.test(String(u || '')) ? u : '#'; }

  var ICON = {
    play: '<svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>',
    pause: '<svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><rect x="6" y="5" width="4" height="14" rx="1"/><rect x="14" y="5" width="4" height="14" rx="1"/></svg>',
    chev: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#5B5F68" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>',
    vid: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>'
  };

  /* ---------- 서버 통신 ---------- */
  function apiGet(q) {
    if (DEMO) return Demo.get(q);
    return fetch(API + '?' + new URLSearchParams(q).toString(), { redirect: 'follow' })
      .then(function (r) { return r.json(); })
      .then(function (j) { if (!j.ok) throw new Error(j.error || '오류가 발생했습니다'); return j; });
  }
  function apiPost(body) {
    if (DEMO) return Demo.post(body);
    return fetch(API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // 사전 요청(CORS preflight) 없이 보내기 위함
      body: JSON.stringify(body),
      redirect: 'follow'
    })
      .then(function (r) { return r.json(); })
      .then(function (j) { if (!j.ok) throw new Error(j.error || '저장하지 못했습니다'); return j; });
  }

  /* ---------- 공통 화면 조각 ---------- */
  function header(right) {
    return '<header class="top">' +
      '<div class="brand"><div class="logo">B4</div><div class="brand-name"><b>B4 Music Lab</b><span>레슨 노트</span></div></div>' +
      right + '</header>' +
      (DEMO ? '<div class="demo-banner">미리보기 모드 · 예시 데이터입니다</div>' : '');
  }
  function message(title, body, extra) {
    $app.innerHTML = '<div class="center-msg">' +
      '<div class="logo">B4</div><h1>' + esc(title) + '</h1><p>' + body + '</p>' + (extra || '') + '</div>';
  }
  function loading() {
    $app.innerHTML = '<div class="center-msg"><div class="spinner" aria-hidden="true"></div><p>불러오는 중…</p></div>';
  }

  /* =========================================================
   * 학생 화면
   * ======================================================= */
  var S = { key: '', data: null, tab: 'home', showAll: false, bpm: null, pct: 100, track: 0, playing: false, loadingAudio: false, loop: false, openRow: null, ctx: 'none' };
  var audio = new Audio();
  audio.preload = 'auto';
  var audioUrls = {};
  var SILENT = 'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEARKwAAIhYAQACABAAZGF0YQAAAAA=';

  function setPitchLock() {
    try { audio.preservesPitch = true; audio.mozPreservesPitch = true; audio.webkitPreservesPitch = true; } catch (e) {}
  }
  setPitchLock();
  audio.addEventListener('timeupdate', updateProgress);
  audio.addEventListener('loadedmetadata', updateProgress);
  audio.addEventListener('ended', function () { S.playing = false; renderStudent(); });

  function startStudent(key) {
    S.key = key;
    loading();
    apiGet({ action: 'student', k: key }).then(function (d) {
      S.data = d;
      var m = d.materials || {};
      S.bpm = m.bpm || null;
      renderStudent();
    }).catch(function (err) {
      message('열 수 없어요', esc(err.message) + '<br>선생님께 받은 링크로 다시 열어주세요.',
        '<button class="primary" style="padding:0 24px" data-act="retry">다시 시도</button>');
    });
  }

  function renderStudent() {
    var d = S.data;
    if (!d) return;
    var html = header('<span class="pill">' + esc(d.name) + ' 님</span>');
    html += S.tab === 'home' ? studentHome(d) : studentLog(d);
    html += '<nav class="tabbar" aria-label="화면 전환">' +
      '<button type="button" data-act="tab" data-tab="home" aria-pressed="' + (S.tab === 'home') + '">이번 주</button>' +
      '<button type="button" data-act="tab" data-tab="log" aria-pressed="' + (S.tab === 'log') + '">내 기록</button></nav>';
    $app.innerHTML = html;
    updateProgress();
  }

  function studentHome(d) {
    var L = d.latest;
    var h = '<div class="stack">';
    h += '<section class="title"><span class="eyebrow">' +
      (L ? esc(L.no + '회차 · ' + L.dateLabel) : '아직 레슨 기록이 없어요') +
      '</span><h1>이번 주 레슨 노트</h1></section>';

    // 다음 수업 날짜
    var N = d.next;
    if (N) {
      h += '<section class="card next-lesson"><span class="label">다음 수업</span>' +
        '<span class="when"><b>' + (N.isToday ? '오늘 · ' : '') + esc(N.label) + '</b>' +
        (N.time ? ' <span class="time">' + esc(N.time) + '</span>' : '') + '</span>' +
        (N.no !== null && N.no !== undefined ? '<span class="pos">' + esc(String(N.no)) + '회차</span>' : '') +
        '</section>';
    }

    if (L) {
      h += '<section class="hero">' +
        '<div class="block"><span class="label">이번 곡</span><span class="song">' + esc(L.song || '—') + '</span>' +
        (L.tech ? '<span class="sub">' + esc(L.tech) + '</span>' : '') + '</div>' +
        '<div class="rule"></div>' +
        '<div class="block"><span class="label">연습 과제</span><span class="sub">' + esc(L.task || '—') + '</span></div>' +
        '</section>';
    }

    // 연습 체크
    var w = d.week;
    h += '<section class="card"><div class="card-head"><h2>이번 주 연습</h2>' +
      '<span class="count">' + w.count + '<small> / 7일</small></span></div><div class="days">';
    w.days.forEach(function (x) {
      var isToday = x.date === d.today;
      h += '<button type="button" class="day' + (x.done ? ' on' : '') + (isToday ? ' today' : '') + '"' +
        ' data-act="practice" data-date="' + x.date + '" aria-pressed="' + x.done + '"' +
        ' aria-label="' + x.label + '요일 ' + (x.done ? '연습함' : '연습 안 함') + '"' + (x.future ? ' disabled' : '') + '>' +
        '<span class="w">' + x.label + '</span><span class="d">' + esc(x.day) + '</span></button>';
    });
    h += '</div><span class="small">연습한 날을 눌러 체크하세요. 다시 누르면 취소됩니다.</span></section>';

    // 자료
    var m = d.materials || { sheets: [], videos: [], audio: [] };
    h += '<section class="card"><div class="card-head"><h2>레슨 자료</h2></div>';
    if (!m.sheets.length && !m.videos.length && !m.audio.length) {
      h += '<p class="empty">아직 등록된 자료가 없어요.</p>';
    }
    h += fileRows(m);
    if (m.audio.length) {
      useCtx(m);
      h += '<div class="subhead"><b>연습 음원</b><span class="small">음정은 그대로, 속도만 조절</span></div>' + playerHTML(m);
    }
    h += '</section>';

    // 레슨 진행
    var p = d.pass;
    if (p) {
      h += '<section class="card"><div class="card-head"><h2>레슨 진행</h2>' +
        '<span class="big">' + p.done + '<small> / ' + p.total + '</small></span></div><div class="bars">';
      for (var i = 0; i < p.total; i++) h += '<i class="' + (i < p.done ? 'on' : '') + '"></i>';
      h += '</div><div class="card-head small"><span>' + p.from + '–' + p.to + '회차</span>' +
        (p.payDue ? '<span class="due">다음 레슨 전 결제</span>' : '<span>결제 완료 · ' + esc(p.paidLabel) + '</span>') +
        '</div></section>';
    }
    return h + '</div>';
  }

  // 지금 플레이어가 어느 회차 음원을 다루는지 맞추기 (바뀌면 정지 후 초기화)
  function useCtx(m) {
    var key = String(m && m.audioRow || 'none');
    if (S.ctx === key) return;
    audio.pause(); S.playing = false; S.track = 0;
    S.ctx = key; S.bpm = (m && m.bpm) || null; S.pct = 100;
  }

  function rateInfo(m) {
    if (m.bpm) {
      var min = Math.round(m.bpm * 0.5), max = Math.round(m.bpm * 1.2);
      var v = Math.min(max, Math.max(min, S.bpm || m.bpm));
      return { mode: 'bpm', value: v, min: min, max: max, rate: v / m.bpm, pct: Math.round(v / m.bpm * 100) };
    }
    var p = Math.min(120, Math.max(50, S.pct));
    return { mode: 'pct', value: p, min: 50, max: 120, rate: p / 100, pct: p };
  }

  function playerHTML(m) {
    var r = rateInfo(m);
    var cur = m.audio[S.track] || m.audio[0];
    var h = '<div class="player">';
    if (m.audio.length > 1) {
      h += '<div class="track-select">';
      m.audio.forEach(function (a, i) {
        h += '<button type="button" class="chip' + (i === S.track ? ' on' : '') + '" data-act="track" data-i="' + i + '" aria-pressed="' + (i === S.track) + '">' + esc(a.label) + '</button>';
      });
      h += '</div>';
    }
    h += '<div class="row"><button type="button" class="play" data-act="play"' + (S.loadingAudio ? ' disabled' : '') +
      ' aria-label="' + (S.playing ? '일시정지' : '재생') + '">' +
      (S.loadingAudio ? '<div class="spinner" style="width:20px;height:20px;border-width:2px"></div>' : (S.playing ? ICON.pause : ICON.play)) +
      '</button><span class="meta"><b>' + esc((m.audioSong ? m.audioSong + ' · ' : '') + (cur ? cur.label : '')) + '</b>' +
      '<span>' + (m.bpm ? '원곡 ' + m.bpm + ' BPM' : '원곡 속도 기준') + '</span></span>' +
      '<button type="button" class="chip' + (S.loop ? ' on' : '') + '" style="flex-shrink:0;white-space:nowrap" data-act="loop" aria-pressed="' + S.loop + '" aria-label="반복 재생">' +
      '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" style="vertical-align:-2px;margin-right:4px"><path d="M17 2l4 4-4 4"/><path d="M3 11V9a3 3 0 0 1 3-3h15"/><path d="M7 22l-4-4 4-4"/><path d="M21 13v2a3 3 0 0 1-3 3H3"/></svg>' +
      '반복</button></div>';
    h += '<div><input class="seek" type="range" min="0" max="1000" value="0" step="1" data-act="seek" aria-label="재생 위치">' +
      '<div class="times"><span data-t="cur">0:00</span><span data-t="dur">' + (audioUrls[tkey(S.track)] ? '' : '--:--') + '</span></div></div>';
    h += '<div class="bpmbox"><button type="button" class="step" data-act="dec" aria-label="느리게">−</button>' +
      '<span class="bpm"><span class="v">' + r.value + '<small>' + (r.mode === 'bpm' ? ' BPM' : '%') + '</small></span>' +
      '<span>' + (r.mode === 'bpm' ? '원곡 대비 ' + r.pct + '%' : '재생 속도') + '</span></span>' +
      '<button type="button" class="step" data-act="inc" aria-label="빠르게">+</button></div>';
    var quick = (m.bpm && m.quick && m.quick.length) ? m.quick.slice() : [];
    if (m.bpm && quick.indexOf(m.bpm) < 0) quick.push(m.bpm);
    if (!m.bpm) quick = [70, 85, 100];
    if (quick.length) {
      h += '<div class="chips"><span>빠른 선택</span>';
      quick.forEach(function (q) {
        var on = q === r.value;
        var label = m.bpm ? (q === m.bpm ? q + ' 원곡' : String(q)) : q + '%';
        h += '<button type="button" class="chip' + (on ? ' on' : '') + '" data-act="quick" data-v="' + q + '" aria-pressed="' + on + '">' + label + '</button>';
      });
      h += '</div>';
    }
    return h + '</div>';
  }

  function curMat() {
    if (S.tab === 'log' && S.openRow) {
      var r = (S.data.sessions || []).filter(function (x) { return x.row === S.openRow; })[0];
      if (r && r.mat) return r.mat;
    }
    return S.data && S.data.materials;
  }

  function applyRate() {
    var m = curMat();
    if (!m) return;
    audio.loop = S.loop;
    setPitchLock();
    audio.playbackRate = rateInfo(m).rate;
  }

  function updateProgress() {
    var seek = $app.querySelector('.seek');
    if (!seek) return;
    var dur = audio.duration;
    if (isFinite(dur) && dur > 0 && audioUrls[tkey(S.track)]) {
      seek.value = Math.round(audio.currentTime / dur * 1000);
      var c = $app.querySelector('[data-t="cur"]'), t = $app.querySelector('[data-t="dur"]');
      if (c) c.textContent = fmtTime(audio.currentTime);
      if (t) t.textContent = fmtTime(dur);
    }
  }

  function tkey(i) { return S.ctx + ':' + i; }
  function loadTrack(i) {
    var key = tkey(i);
    if (audioUrls[key]) return Promise.resolve(audioUrls[key]);
    return apiGet({ action: 'audio', k: S.key, i: i, row: S.ctx === 'none' ? '' : S.ctx }).then(function (j) {
      var bin = atob(j.data), arr = new Uint8Array(bin.length);
      for (var n = 0; n < bin.length; n++) arr[n] = bin.charCodeAt(n);
      audioUrls[key] = URL.createObjectURL(new Blob([arr], { type: j.mime || 'audio/mpeg' }));
      return audioUrls[key];
    });
  }

  function togglePlay() {
    if (S.playing) { audio.pause(); S.playing = false; renderStudent(); return; }
    if (audioUrls[tkey(S.track)]) {
      if (audio.src !== audioUrls[tkey(S.track)]) audio.src = audioUrls[tkey(S.track)];
      applyRate();
      audio.play().then(function () { S.playing = true; renderStudent(); })
        .catch(function () { toast('재생 버튼을 한 번 더 눌러주세요'); });
      return;
    }
    // iOS: 사용자가 누른 순간에 재생을 한 번 시작해 둬야 이후 자동 재생이 허용됨
    audio.src = SILENT;
    audio.play().catch(function () {});
    S.loadingAudio = true; renderStudent();
    loadTrack(S.track).then(function (url) {
      audio.src = url;
      applyRate();
      return audio.play().then(function () { S.playing = true; })
        .catch(function () { S.playing = false; toast('준비됐어요. 재생 버튼을 눌러주세요'); });
    }).catch(function (err) {
      toast(err.message || '음원을 불러오지 못했어요');
    }).then(function () { S.loadingAudio = false; renderStudent(); });
  }

  function changeRate(delta, absolute) {
    var m = curMat(), r = rateInfo(m);
    var v = absolute !== undefined ? absolute : r.value + delta * (r.mode === 'bpm' ? 1 : 5);
    v = Math.min(r.max, Math.max(r.min, v));
    if (r.mode === 'bpm') S.bpm = v; else S.pct = v;
    applyRate();
    renderStudent();
  }

  function togglePractice(date) {
    var w = S.data.week;
    var day = w.days.filter(function (x) { return x.date === date; })[0];
    if (!day || day.future) return;
    day.done = !day.done;
    w.count = w.days.filter(function (x) { return x.done; }).length;
    renderStudent();
    apiPost({ action: 'practice', k: S.key, date: date, done: day.done }).catch(function (err) {
      day.done = !day.done;
      w.count = w.days.filter(function (x) { return x.done; }).length;
      renderStudent();
      toast(err.message || '저장하지 못했어요. 다시 눌러주세요');
    });
  }

  // 악보·영상 칸의 자료 → 파일 형식(확장자)만 표시
  var VIDEO_EXT = ['YouTube', 'MP4', 'MOV', 'M4V', 'AVI', 'MKV', 'WEBM'];
  function fileRows(m) {
    var h = '';
    (m.sheets || []).concat(m.videos || []).forEach(function (x) {
      var ext = x.ext || '';
      var isVid = VIDEO_EXT.indexOf(ext) >= 0;
      var icon = ext === 'YouTube' ? ICON.vid : (ext ? esc(ext) : '링크');
      var cls = 'ic ' + (isVid ? 'vid' : 'pdf') + (ext.length > 3 && ext !== 'YouTube' ? ' long' : '');
      h += '<a class="res" href="' + esc(safeUrl(x.url)) + '" target="_blank" rel="noopener">' +
        '<span class="' + cls + '">' + icon + '</span><span class="txt"><b>' + esc(x.label) + '</b>' +
        (ext ? '<span class="small">' + esc(ext) + '</span>' : '') + '</span>' + ICON.chev + '</a>';
    });
    return h;
  }

  function rowDetail(r) {
    var h = '<div class="row-detail">';
    if (r.song) h += '<div class="kv"><span>곡</span><b>' + esc(r.song) + '</b></div>';
    if (r.tech) h += '<div class="kv"><span>테크닉·이론</span><b>' + esc(r.tech) + '</b></div>';
    if (r.task) h += '<div class="kv"><span>연습 과제</span><b>' + esc(r.task) + '</b></div>';
    var m = r.mat;
    if (!m) {
      h += '<p class="empty">이 회차에 등록된 자료가 없어요.</p>';
    } else {
      h += fileRows(m);
      if (m.audio.length) { useCtx(m); h += playerHTML(m); }
    }
    return h + '</div>';
  }

  function studentLog(d) {
    var st = d.stats || {};
    var h = '<div class="stack">';
    h += '<section class="title"><span class="eyebrow">' + esc(st.since ? st.since + '부터' : '') + '</span><h1>내 기록</h1></section>';
    h += '<div class="stats"><div class="stat dark"><span class="small">총 레슨</span><span class="num">' + (st.total || 0) + '<small> 회</small></span></div>' +
      '<div class="stat"><span class="small">레슨 기간</span><span class="num">' +
      (st.years >= 1 ? st.years + '<small> 년+</small>' : '1<small> 년 미만</small>') + '</span></div></div>';

    var rows = d.sessions || [];
    var shown = S.showAll ? rows : rows.slice(0, 5);
    h += '<section class="card" style="gap:0"><div class="card-head" style="margin-bottom:8px"><h2>레슨 기록</h2><span class="small">최근순</span></div>';
    shown.forEach(function (r) {
      if (r.isBreak) { h += '<div class="break-row">휴강</div>'; return; }
      // 납부 완료 묶음은 색(주황·초록·파랑 순환), 납부 전 묶음은 회색 점선
      var tone = r.trial ? 'trial' : (r.block === undefined ? 'none' : (r.paid ? 'b' + (r.block % 3) : 'unpaid'));
      var tag = r.trial ? '<span class="pos trial">시범</span>' :
        (r.pos ? '<span class="pos">' + r.pos + '/4</span>' + (r.paid ? '' : '<span class="pos unpaid">납부 전</span>') : '');
      var open = S.openRow === r.row;
      h += '<button type="button" class="list-row tone-' + tone + (open ? ' open' : '') + '" data-act="openrow" data-row="' + r.row + '" aria-expanded="' + open + '">' +
        '<span class="no"><b>' + r.no + '</b><span>회차</span></span>' +
        '<span class="txt"><span class="when">' + esc(r.dateLabel) + ' ' + tag + (r.mat ? '<span class="pos has-mat">자료</span>' : '') + '</span>' +
        '<span class="what">' + esc(r.content || '—') + '</span>' +
        (r.pay ? '<span class="paid">수강료 납부 · ' + esc(r.pay.date) + (r.pay.fee ? ' · ' + esc(r.pay.fee) : '') + '</span>' : '') +
        '</span><span class="chev" aria-hidden="true">' + (open ? '▴' : '▾') + '</span></button>';
      if (open) h += rowDetail(r);
    });
    if (rows.length > 5) {
      h += '<button type="button" class="btn-soft" style="margin-top:8px" data-act="more">' +
        (S.showAll ? '접기' : '이전 기록 더 보기 (전체 ' + (st.total || rows.length) + '회)') + '</button>';
    }
    h += '</section>';

    var pays = d.payments || [];
    if (pays.length) {
      var shownPays = S.showAllPays ? pays : pays.slice(0, 5);
      h += '<section class="card" style="gap:0"><div class="card-head" style="margin-bottom:8px"><h2>수강료 납부 내역</h2><span class="small">' + pays.length + '회</span></div>';
      shownPays.forEach(function (p, i) {
        h += '<div class="pay-row tone-b' + ((p.block !== undefined ? p.block : pays.length - 1 - i) % 3) + '"><span class="txt"><span class="when">' + esc(p.date) + '</span>' +
          '<span class="what">' + p.from + '–' + p.to + '회차</span></span><span class="fee">' + esc(p.fee || '') + '</span></div>';
      });
      if (pays.length > 5) {
        h += '<button type="button" class="btn-soft" style="margin-top:8px" data-act="morepays">' + (S.showAllPays ? '접기' : '전체 보기') + '</button>';
      }
      h += '</section>';
    }

    var rep = d.repertoire || [];
    if (rep.length) {
      h += '<section class="card" style="gap:0"><div class="card-head" style="margin-bottom:8px"><h2>레퍼토리</h2><span class="small">' + rep.length + '곡</span></div>';
      rep.forEach(function (r) {
        h += '<div class="rep-row"><span class="when">' + esc(r.when) + '</span><span><b>' + esc(r.title) + '</b><span class="small">' + esc(r.meta) + '</span></span></div>';
      });
      h += '</section>';
    }
    return h + '</div>';
  }

  /* =========================================================
   * 선생님 화면
   * ======================================================= */
  var T = { key: '', data: null, scope: 'today', open: null, busy: null };

  function startTeacher(key) {
    T.key = key;
    loadTeacher();
  }

  function loadTeacher(keepScreen) {
    if (!keepScreen) loading();
    return apiGet({ action: 'admin', admin: T.key, scope: T.scope }).then(function (d) {
      T.data = d; renderTeacher();
    }).catch(function (err) {
      message('열 수 없어요', esc(err.message), '<button class="primary" style="padding:0 24px" data-act="retry">다시 시도</button>');
    });
  }

  function renderTeacher() {
    var d = T.data;
    var list = d.students || [];
    var attended = list.filter(function (s) { return s.attended; }).length;
    var h = header('<span class="pill dark">선생님</span>') + '<div class="stack">';
    h += '<section class="title"><span class="eyebrow">' + esc(d.todayLabel) + '</span><h1>' +
      (T.scope === 'today' ? '오늘 레슨 <span class="num" style="color:#5B5F68">' + attended + '/' + list.length + '</span>' : '전체 학생 <span class="num" style="color:#5B5F68">' + list.length + '</span>') +
      '</h1></section>';
    h += '<div class="seg" role="group" aria-label="보기">' +
      '<button type="button" data-act="scope" data-v="today" aria-pressed="' + (T.scope === 'today') + '">오늘 레슨</button>' +
      '<button type="button" data-act="scope" data-v="all" aria-pressed="' + (T.scope === 'all') + '">전체 학생</button></div>';

    if (!list.length) {
      h += '<section class="card"><p class="empty">' + (T.scope === 'today'
        ? '오늘 요일로 등록된 학생이 없어요. 보강 레슨은 "전체 학생"에서 출석 처리하세요.'
        : '[학생] 탭에 사용 체크된 학생이 없어요.') + '</p></section>';
    }
    list.forEach(function (s) { h += teacherCard(s); });
    $app.innerHTML = h + '</div>';
  }

  // 학생 카드 이름 아래: 다음 수업 날짜·요일·시간·회차
  function nextLine(s) {
    var U = s.upcoming;
    if (!U) return [s.time, s.next ? s.next + '회차' : ''].filter(Boolean).join(' · ');
    return ['다음 ' + (U.isToday ? '오늘 ' : '') + U.label, U.time, (U.no !== null && U.no !== undefined) ? U.no + '회차' : '']
      .filter(Boolean).join(' · ');
  }

  function teacherCard(s) {
    if (s.error) {
      return '<section class="card"><b>' + esc(s.name) + '</b><p class="empty">시트를 읽지 못했어요: ' + esc(s.error) + '</p></section>';
    }
    var busy = T.busy === s.key;
    var h = '<section class="card stu"><div class="stu-top"><div class="who"><span class="avatar">' + esc(String(s.name).charAt(0)) + '</span>' +
      '<span class="txt"><b>' + esc(s.name) + '</b><span class="small">' + esc(nextLine(s)) + '</span></span></div>' +
      '<button type="button" class="attend' + (s.attended ? ' on' : '') + '" data-act="attend" data-key="' + esc(s.key) + '"' +
      (s.attended || busy ? ' disabled' : '') + ' aria-pressed="' + !!s.attended + '">' + (busy ? '처리 중…' : (s.attended ? '출석 완료' : '출석')) + '</button></div>';
    h += '<div class="tags"><span class="tag">최근 7일 연습 ' + s.practice + '/7</span>' +
      (s.pass ? '<span class="tag">레슨 진행 ' + s.pass.done + '/' + s.pass.total + '</span>' : '') +
      (s.pass && s.pass.payDue ? '<span class="due">결제 필요</span>' : '') + '</div>';
    h += '<div class="links"><button type="button" class="linkbtn" data-act="form" data-key="' + esc(s.key) + '">' + (T.open === s.key ? '입력 닫기' : '레슨 기록 입력') + '</button>' +
      (s.link ? '<a class="linkbtn" href="' + esc(safeUrl(s.link)) + '" target="_blank" rel="noopener">학생 화면 보기</a>' +
        '<button type="button" class="linkbtn" data-act="copy" data-link="' + esc(s.link) + '">학생 링크 복사</button>' : '') + '</div>';
    if (T.open === s.key) h += recordForm(s);
    return h + '</section>';
  }

  function recordForm(s) {
    var L = s.latest || {};
    var warn = s.attended ? '' : '<p class="small" style="margin:0">오늘 출석 전이라 ' + (L.no !== undefined ? L.no + '회차' : '최근 회차') + '에 저장됩니다. 오늘 레슨이면 먼저 출석을 누르세요.</p>';
    function field(name, label, val, type, hint, area) {
      var input = area
        ? '<textarea name="' + name + '" rows="2">' + esc(val) + '</textarea>'
        : '<input name="' + name + '" type="' + (type || 'text') + '" value="' + esc(val) + '"' + (type === 'number' ? ' inputmode="numeric"' : '') + (name === 'quick' ? ' placeholder="예: 60,72"' : '') + (name === 'fee' ? ' placeholder="예: 130000"' : '') + '>';
      return '<label>' + label + (hint ? ' <span class="hint">' + hint + '</span>' : '') + input + '</label>';
    }
    return '<form class="form" data-key="' + esc(s.key) + '">' + warn +
      field('song', '곡', L.song) +
      field('tech', '테크닉 · 이론', L.tech) +
      field('task', '연습 과제', L.task, '', '', true) +
      field('sheets', '악보 링크', L.sheets, '', '한 줄에 하나 · "제목 | 주소" 가능', true) +
      field('videos', '영상 링크', L.videos, '', '한 줄에 하나', true) +
      field('audio', '음원 링크', L.audio, '', '드라이브 MP3 · 한 줄에 하나', true) +
      '<div class="two">' + field('bpm', '원곡 BPM', L.bpm, 'number') + field('quick', '빠른 선택 BPM', L.quick, 'text') + '</div>' +
      '<div class="two"><label class="check"><input type="checkbox" name="paid"> 이번 회차에 결제 받음</label>' +
      field('fee', '수강료(원)', '', 'number') + '</div>' +
      '<button type="submit" class="primary"' + (T.busy === s.key ? ' disabled' : '') + '>시트에 저장</button></form>';
  }

  function doAttend(key) {
    T.busy = key; renderTeacher();
    apiPost({ action: 'attend', admin: T.key, k: key }).then(function (j) {
      toast(j.already ? '이미 오늘 출석 처리됐어요' : j.no + '회차 출석 처리했어요');
      T.busy = null;
      return loadTeacher(true);
    }).catch(function (err) { T.busy = null; renderTeacher(); toast(err.message); });
  }

  function doRecord(form) {
    var key = form.getAttribute('data-key');
    var f = form.elements;
    var body = {
      action: 'record', admin: T.key, k: key,
      song: f.song.value, tech: f.tech.value, task: f.task.value,
      sheets: f.sheets.value, videos: f.videos.value, audio: f.audio.value,
      bpm: f.bpm.value, quick: f.quick.value, paid: f.paid.checked, fee: f.fee.value
    };
    T.busy = key; renderTeacher();
    apiPost(body).then(function (j) {
      toast(j.no + '회차에 저장했어요');
      T.busy = null; T.open = null;
      return loadTeacher(true);
    }).catch(function (err) { T.busy = null; renderTeacher(); toast(err.message); });
  }

  /* =========================================================
   * 이벤트
   * ======================================================= */
  $app.addEventListener('click', function (e) {
    var b = e.target.closest('[data-act]');
    if (!b || b.disabled) return;
    var act = b.getAttribute('data-act');
    switch (act) {
      case 'retry': location.reload(); break;
      case 'tab': S.tab = b.getAttribute('data-tab'); S.openRow = null; renderStudent(); window.scrollTo(0, 0); break;
      case 'more': S.showAll = !S.showAll; renderStudent(); break;
      case 'morepays': S.showAllPays = !S.showAllPays; renderStudent(); break;
      case 'openrow':
        var row = Number(b.getAttribute('data-row'));
        S.openRow = S.openRow === row ? null : row;
        if (!S.openRow) { audio.pause(); S.playing = false; }
        renderStudent();
        break;
      case 'practice': togglePractice(b.getAttribute('data-date')); break;
      case 'play': togglePlay(); break;
      case 'loop': S.loop = !S.loop; audio.loop = S.loop; renderStudent(); break;
      case 'dec': changeRate(-1); break;
      case 'inc': changeRate(1); break;
      case 'quick': changeRate(0, Number(b.getAttribute('data-v'))); break;
      case 'track':
        var i = Number(b.getAttribute('data-i'));
        if (i !== S.track) { audio.pause(); S.playing = false; S.track = i; if (audioUrls[tkey(i)]) audio.src = audioUrls[tkey(i)]; renderStudent(); }
        break;
      case 'scope': T.scope = b.getAttribute('data-v'); T.open = null; loadTeacher(); break;
      case 'attend': doAttend(b.getAttribute('data-key')); break;
      case 'form': var k = b.getAttribute('data-key'); T.open = T.open === k ? null : k; renderTeacher(); break;
      case 'copy':
        var link = b.getAttribute('data-link');
        if (navigator.clipboard) navigator.clipboard.writeText(link).then(function () { toast('학생 링크를 복사했어요'); }, function () { toast(link); });
        else toast(link);
        break;
      case 'demo': location.href = '?demo=1&view=' + b.getAttribute('data-v'); break;
      case 'savekey':
        var inp = document.getElementById('keyinput');
        var v = (inp && inp.value || '').trim();
        var ad = v.match(/[?&]admin=([A-Za-z0-9]+)/); // 선생님 링크를 붙여넣은 경우
        if (ad) { store.set('b4admin', ad[1]); location.href = './?admin=' + encodeURIComponent(ad[1]); return; }
        var m = v.match(/[?&]k=([A-Za-z0-9]+)/);
        var key = m ? m[1] : v;
        if (key.length < 12) { toast('링크를 다시 확인해주세요'); return; }
        store.set('b4k', key);
        location.href = './?k=' + encodeURIComponent(key);
        break;
    }
  });

  $app.addEventListener('input', function (e) {
    if (e.target.classList.contains('seek') && isFinite(audio.duration) && audioUrls[tkey(S.track)]) {
      audio.currentTime = e.target.value / 1000 * audio.duration;
    }
  });

  $app.addEventListener('submit', function (e) {
    var form = e.target.closest('form.form');
    if (!form) return;
    e.preventDefault();
    doRecord(form);
  });

  /* =========================================================
   * 미리보기(데모) 데이터 — config.js 의 API_URL 이 비어 있을 때만 사용
   * ======================================================= */
  var Demo = (function () {
    function pad(n) { return (n < 10 ? '0' : '') + n; }
    function ds(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
    var now = new Date(), todayS = ds(now);
    var dow = (now.getDay() + 6) % 7; // 0=월
    var practice = {};
    function week() {
      var labels = ['월', '화', '수', '목', '금', '토', '일'], days = [];
      for (var i = 0; i < 7; i++) {
        var d = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow + i);
        var s = ds(d);
        days.push({ date: s, label: labels[i], day: String(d.getDate()), done: !!practice[s], future: s > todayS });
      }
      return { days: days, count: days.filter(function (x) { return x.done; }).length };
    }
    // 예시: 이번 주 첫날·둘째 날 연습한 것으로 표시
    (function () { var w = week().days; if (!w[0].future) practice[w[0].date] = true; if (!w[1].future) practice[w[1].date] = true; })();

    var sessionsRaw = [
      { no: 107, dateLabel: '2026.02.25 (수)', content: '밤양갱 코러스부터' },
      { no: 106, dateLabel: '2026.02.11 (수)', content: '밤양갱' },
      { no: 105, dateLabel: '2026.01.28 (수)', content: "We're Going to Be Friends / Stand by Me · Settling the Spirit 쳐보기" },
      { no: 104, dateLabel: '2026.01.21 (수)', content: "Settling the Spirit 복습 / We're Going to Be Friends" },
      { no: 103, dateLabel: '2026.01.14 (수)', content: 'Settling the Spirit 복습 / Sea of Love' },
      { no: 102, dateLabel: '2026.01.07 (수)', content: 'Settling the Spirit 마지막 / Baby One More Time · Essentials Fingerstyles' },
      { no: 101, dateLabel: '2025.12.31 (수)', content: 'Settling the Spirit / 작은 자유' },
      { no: 100, dateLabel: '2025.12.24 (수)', content: 'Settling the Spirit / Santa Baby' },
      { no: 99, dateLabel: '2025.12.17 (수)', content: 'Settling the Spirit / Last Christmas' },
      { no: 98, dateLabel: '2025.12.03 (수)', content: 'Settling the Spirit' }
    ];
    // 미리보기: 105회차에 결제, 101회차에 결제 → 4회 묶음 표시
    var sessions = sessionsRaw.map(function (r, i) {
      var x = Object.assign({}, r);
      x.row = 200 - i; x.song = r.content; x.tech = ''; x.task = '';
      if (r.no === 107) x.mat = { sheets: [{ label: '밤양갱_TAB', url: '#', ext: 'PDF' }, { label: '밤양갱_코드표', url: '#', ext: 'PDF' }], videos: [], audio: [{ index: 0, label: '메트로놈 예시' }], audioSong: '밤양갱', audioRow: x.row, bpm: 84, quick: [60, 72] };
      if (r.no === 103) x.mat = { sheets: [{ label: 'Settling the Spirit', url: '#', ext: 'PDF' }], videos: [{ label: 'Sea of Love', url: '#', ext: 'YouTube' }], audio: [{ index: 0, label: '메트로놈 예시' }], audioSong: 'Settling the Spirit', audioRow: x.row, bpm: 70, quick: [] };
      if (r.no >= 105) { x.block = 2; x.pos = r.no - 104; x.paid = true; }
      else if (r.no >= 101) { x.block = 1; x.pos = r.no - 100; x.paid = true; }
      else { x.block = 0; x.pos = r.no - 96; x.paid = false; }
      if (r.no === 105) x.pay = { date: '2026.03.04', fee: '130,000원' };
      if (r.no === 101) x.pay = { date: '2026.01.21', fee: '130,000원' };
      return x;
    });
    var student = {
      ok: true, name: 'ㄱ', today: todayS,
      latest: { no: 107, dateLabel: '2026.02.25 (수)', song: '밤양갱', tech: '코러스부터', task: '핑거스타일 기본 연습 구간' },
      materials: {
        sheets: [{ label: '밤양갱_TAB', url: '#', ext: 'PDF' }, { label: '밤양갱_코드표', url: '#', ext: 'PDF' }],
        videos: [{ label: '핑거스타일 기본 연습', url: '#', ext: 'MP4' }],
        audio: [{ index: 0, label: '메트로놈 예시' }, { index: 1, label: '반주만' }],
        audioSong: '밤양갱', audioRow: 200, bpm: 84, quick: [60, 72]
      },
      pass: { done: 3, total: 4, from: 105, to: 108, paidLabel: '3/4', payDue: false },
      next: { date: '', label: '11/18 (수)', time: '19:00', no: 108, isToday: false },
      sessions: sessions,
      payments: [
        { date: '2026.03.04', fee: '130,000원', from: 105, to: 108, block: 2 },
        { date: '2026.01.21', fee: '130,000원', from: 101, to: 104, block: 1 }
      ],
      repertoire: [
        { when: '26.02', title: '밤양갱', meta: '2회 레슨 · 마지막 107회차' },
        { when: '26.01', title: "We're Going to Be Friends", meta: '2회 레슨 · 마지막 105회차' },
        { when: '26.01', title: 'Stand by Me', meta: '105회차' },
        { when: '25.11', title: 'Settling the Spirit', meta: '8회 레슨 · 마지막 105회차' },
        { when: '25.10', title: 'Simon and Garfunkel - 59th', meta: '3회 레슨 · 마지막 97회차' },
        { when: '25.09', title: '가을아침', meta: '2회 레슨 · 마지막 94회차' }
      ],
      stats: { total: 108, since: '2023.01.07', years: 3 }
    };
    var teacherStudents = [
      { key: 'demo-g', name: 'ㄱ', time: '14:00', next: 108, attended: false, practice: 2, upcoming: { label: '11/16 (월)', time: '14:00', no: 108, isToday: false },
        pass: { done: 3, total: 4, payDue: false }, link: '?demo=1&view=student',
        latest: { no: 107, song: '밤양갱', tech: '코러스부터', task: '핑거스타일 기본 연습 구간', sheets: '', videos: '', audio: '', bpm: 84, quick: '60,72' } },
      { key: 'demo-b', name: '학생 B (예시)', time: '16:00', next: 24, attended: true, practice: 5, upcoming: { label: '11/21 (토)', time: '16:00', no: 25, isToday: false },
        pass: { done: 4, total: 4, payDue: true }, link: '',
        latest: { no: 24, song: '[곡]', tech: '', task: '', sheets: '', videos: '', audio: '', bpm: '', quick: '' } },
      { key: 'demo-c', name: '학생 C (예시)', time: '19:00', next: 9, attended: false, practice: 0, upcoming: { label: '11/14 (토)', time: '19:00', no: 9, isToday: true },
        pass: { done: 1, total: 4, payDue: false }, link: '',
        latest: { no: 8, song: '[곡]', tech: '', task: '', sheets: '', videos: '', audio: '', bpm: '', quick: '' } }
    ];

    // 미리보기용 메트로놈 음원 (84 BPM, 4/4, 8마디) — 속도 조절을 바로 들어볼 수 있게
    function metronomeWav(bpm, beats) {
      var sr = 22050, spb = 60 / bpm, len = Math.floor(sr * spb * beats);
      var data = new Int16Array(len);
      for (var b = 0; b < beats; b++) {
        var start = Math.floor(b * spb * sr), f = b % 4 === 0 ? 1600 : 1000;
        for (var i = 0; i < sr * 0.04 && start + i < len; i++) {
          data[start + i] = Math.sin(2 * Math.PI * f * i / sr) * Math.exp(-i / (sr * 0.012)) * 20000;
        }
      }
      var buf = new ArrayBuffer(44 + len * 2), v = new DataView(buf);
      function w(o, s) { for (var i = 0; i < s.length; i++) v.setUint8(o + i, s.charCodeAt(i)); }
      w(0, 'RIFF'); v.setUint32(4, 36 + len * 2, true); w(8, 'WAVE'); w(12, 'fmt ');
      v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true);
      v.setUint32(24, sr, true); v.setUint32(28, sr * 2, true); v.setUint16(32, 2, true); v.setUint16(34, 16, true);
      w(36, 'data'); v.setUint32(40, len * 2, true);
      for (var k = 0; k < len; k++) v.setInt16(44 + k * 2, data[k], true);
      var bytes = new Uint8Array(buf), bin = '';
      for (var c = 0; c < bytes.length; c += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(c, c + 0x8000));
      return btoa(bin);
    }

    function later(v) { return new Promise(function (res) { setTimeout(function () { res(v); }, 250); }); }
    return {
      get: function (q) {
        if (q.action === 'student') { student.week = week(); return later(JSON.parse(JSON.stringify(student))); }
        if (q.action === 'audio') return later({ ok: true, mime: 'audio/wav', data: metronomeWav(84, 32) });
        if (q.action === 'admin') {
          return later({ ok: true, todayLabel: (now.getMonth() + 1) + '월 ' + now.getDate() + '일 (' + '일월화수목금토'.charAt(now.getDay()) + ')',
            scope: q.scope, students: JSON.parse(JSON.stringify(teacherStudents)) });
        }
        return Promise.reject(new Error('알 수 없는 요청'));
      },
      post: function (b) {
        if (b.action === 'practice') { practice[b.date] = b.done; return later({ ok: true }); }
        var s = teacherStudents.filter(function (x) { return x.key === b.k; })[0];
        if (b.action === 'attend' && s) { s.attended = true; s.pass.done = Math.min(4, s.pass.done + 1); s.pass.payDue = s.pass.done >= 4; return later({ ok: true, no: s.next }); }
        if (b.action === 'record' && s) {
          ['song', 'tech', 'task', 'sheets', 'videos', 'audio', 'bpm', 'quick'].forEach(function (f) { s.latest[f] = b[f]; });
          return later({ ok: true, no: s.attended ? s.next : s.latest.no });
        }
        return Promise.reject(new Error('알 수 없는 요청'));
      }
    };
  })();

  /* =========================================================
   * 시작
   * ======================================================= */
  function boot() {
    var k = params.get('k'), admin = params.get('admin');
    if (DEMO) {
      var view = params.get('view');
      if (view === 'student') return startStudent('demo');
      if (view === 'teacher') return startTeacher('demo');
      return message('B4 레슨 노트', '미리보기 모드입니다.<br>config.js 에 데이터 주소를 넣으면 실제 시트와 연결됩니다.',
        '<div class="keyform"><button class="primary" data-act="demo" data-v="student">학생 화면 보기</button>' +
        '<button class="btn-soft" data-act="demo" data-v="teacher">선생님 화면 보기</button></div>');
    }
    if (admin) { store.set('b4admin', admin); return startTeacher(admin); }
    if (k) { store.set('b4k', k); return startStudent(k); }
    // 홈 화면 아이콘으로 열린 경우: 저장된 키 사용
    var savedAdmin = store.get('b4admin'), savedK = store.get('b4k');
    if (savedAdmin) return startTeacher(savedAdmin);
    if (savedK) return startStudent(savedK);
    message('링크로 열어주세요', '받은 링크(학생 링크 또는 선생님 링크)를<br>아래에 붙여넣으면 다음부터는 바로 열립니다.',
      '<div class="keyform"><input id="keyinput" type="url" placeholder="받은 링크 붙여넣기" aria-label="받은 링크">' +
      '<button class="primary" data-act="savekey">열기</button></div>');
  }

  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); });
  }
  boot();
})();
