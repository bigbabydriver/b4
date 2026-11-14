/**
 * B4 레슨 노트 — 데이터 창구 (Apps Script)
 * ------------------------------------------------------------
 * 위치: "B4 레슨 마스터" 시트에 연결된 Apps Script 프로젝트
 *       (같은 프로젝트에 시트정리_스크립트.gs 도 함께 둠)
 *
 * 역할: PWA(학생·선생님 화면)가 보내는 요청을 받아 구글 시트를 읽고 씀
 *   GET  ?action=student&k=학생키            → 학생 화면 데이터
 *   GET  ?action=audio&k=학생키&i=0          → 연습 음원(base64)
 *   GET  ?action=admin&admin=관리자키&scope=today|all → 선생님 화면 데이터
 *   POST {action:'practice', k, date, done}  → 연습 체크 저장
 *   POST {action:'attend', admin, k}         → 출석 (새 회차 행 추가)
 *   POST {action:'record', admin, k, ...}    → 최근 회차 레슨 내용 저장
 *
 * [학생] 탭 열 구성
 *   A 사용(체크) | B 이름 | C 시트 주소 | D 레슨 요일(예: 금 / 월,목) | E 레슨 시간 | F 링크 키 | G 학생 링크
 */

var TZ = 'Asia/Seoul';
var STUDENT_TAB = '학생';
var LOG_TAB = '레슨기록';
var PRACTICE_TAB = '연습기록';
var PASS_SIZE = 4;                 // 4회 단위 결제
var CLEAN_FOLDER = 'B4 학생 시트 정리본'; // 정리본 시트가 들어 있는 드라이브 폴더
var MAX_AUDIO_BYTES = 20 * 1024 * 1024;
var DOW = ['일', '월', '화', '수', '목', '금', '토'];

// 레슨기록 열 번호 (1부터)
var C = { DONE: 1, NO: 2, DATE: 3, PAY: 4, SONG: 5, TECH: 6, TASK: 7, REF: 8,
          SHEETS: 9, VIDEOS: 10, AUDIO: 11, BPM: 12, QUICK: 13, NOTE: 14, FEE: 15 };

/* =============================================================
 * 메뉴
 * =========================================================== */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  if (typeof addCleanupMenu_ === 'function') addCleanupMenu_(ui);
  ui.createMenu('B4 앱')
    .addItem('1. 학생 탭 만들기', 'setupStudentTab')
    .addItem('2. 앱 주소 등록', 'setAppUrl')
    .addItem('3. 학생 링크 발급 (빈 칸만)', 'issueLinks')
    .addSeparator()
    .addItem('선택한 학생 링크 재발급', 'reissueSelectedLink')
    .addItem('선생님(관리자) 링크 보기', 'showAdminLink')
    .addItem('선생님 링크 재발급', 'reissueAdminLink')
    .addToUi();
}

function setupStudentTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STUDENT_TAB) || ss.insertSheet(STUDENT_TAB);
  var headers = ['사용', '이름', '시트 주소', '레슨 요일', '레슨 시간', '링크 키', '학생 링크'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  // 비어 있으면 1순위: 드라이브 'B4 학생 시트 정리본' 폴더, 2순위: [학생목록]에서 체크한 시트
  if (sh.getLastRow() < 2) {
    var folders = DriveApp.getFoldersByName(CLEAN_FOLDER);
    if (folders.hasNext()) {
      var files = folders.next().getFilesByType(MimeType.GOOGLE_SHEETS), rowsF = [];
      while (files.hasNext()) {
        var f = files.next();
        rowsF.push([true, f.getName().replace(/^_+/, ''), f.getUrl(), '', '', '', '']);
      }
      rowsF.sort(function (a, b) { return a[1] < b[1] ? -1 : 1; });
      if (rowsF.length) sh.getRange(2, 1, rowsF.length, headers.length).setValues(rowsF);
    }
  }
  if (sh.getLastRow() < 2) {
    var list = ss.getSheetByName('학생목록');
    if (list && list.getLastRow() > 1) {
      var src = list.getRange(2, 1, list.getLastRow() - 1, 3).getValues()
        .filter(function (r) { return r[0] === true; })
        .map(function (r) { return [true, String(r[1]).replace(/^_+/, ''), r[2], '', '', '', '']; });
      if (src.length) sh.getRange(2, 1, src.length, headers.length).setValues(src);
    }
  }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).insertCheckboxes();
  sh.autoResizeColumns(1, headers.length);
  SpreadsheetApp.getUi().alert('[학생] 탭을 준비했습니다.\nD열(레슨 요일)과 E열(시간)을 채운 뒤, "앱 주소 등록" → "학생 링크 발급" 순서로 진행하세요.');
}

function setAppUrl() {
  var ui = SpreadsheetApp.getUi();
  var cur = props_().getProperty('APP_URL') || '';
  var res = ui.prompt('앱 주소 등록', '앱(PWA) 주소를 입력하세요. 예: https://계정이름.github.io/b4/\n현재: ' + (cur || '(없음)'), ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var url = res.getResponseText().trim();
  if (!/^https:\/\//.test(url)) { ui.alert('https:// 로 시작하는 주소를 입력하세요.'); return; }
  if (!/\/$/.test(url) && !/\.html$/.test(url)) url += '/';
  props_().setProperty('APP_URL', url);
  refreshLinkColumn_();
  ui.alert('등록했습니다: ' + url);
}

function issueLinks() {
  var sh = studentSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) { SpreadsheetApp.getUi().alert('[학생] 탭에 학생이 없습니다.'); return; }
  var vals = sh.getRange(2, 1, n, 7).getValues();
  var count = 0;
  vals.forEach(function (r) {
    if (r[1] && r[2] && !r[5]) { r[5] = randomKey_(); count++; }
  });
  sh.getRange(2, 1, n, 7).setValues(vals);
  refreshLinkColumn_();
  ensureAdminKey_();
  SpreadsheetApp.getUi().alert(count + '명의 링크를 새로 발급했습니다.\nG열의 학생 링크를 각 학생에게 개별로 보내세요.');
}

function reissueSelectedLink() {
  var sh = studentSheet_();
  var row = sh.getActiveRange() ? sh.getActiveRange().getRow() : 0;
  var ui = SpreadsheetApp.getUi();
  if (SpreadsheetApp.getActiveSheet().getName() !== STUDENT_TAB || row < 2) {
    ui.alert('[학생] 탭에서 재발급할 학생의 행을 선택한 뒤 실행하세요.'); return;
  }
  var name = sh.getRange(row, 2).getValue();
  var ok = ui.alert(name + ' 학생의 링크를 재발급합니다.\n기존 링크는 바로 쓸 수 없게 됩니다. 진행할까요?', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  sh.getRange(row, 6).setValue(randomKey_());
  refreshLinkColumn_();
  ui.alert('재발급했습니다. G열의 새 링크를 학생에게 다시 보내세요.');
}

function showAdminLink() {
  ensureAdminKey_();
  var url = props_().getProperty('APP_URL');
  var key = props_().getProperty('ADMIN_KEY');
  SpreadsheetApp.getUi().alert(url
    ? '선생님 링크 (다른 사람에게 보내지 마세요)\n\n' + url + '?admin=' + key
    : '먼저 "앱 주소 등록"을 실행하세요.');
}

function reissueAdminLink() {
  props_().setProperty('ADMIN_KEY', randomKey_() + randomKey_());
  showAdminLink();
}

/* =============================================================
 * 요청 처리
 * =========================================================== */
function doGet(e) {
  try {
    var p = e.parameter || {};
    if (p.action === 'student') return json_(studentPayload_(requireStudent_(p.k)));
    if (p.action === 'audio') return json_(audioPayload_(requireStudent_(p.k), parseInt(p.i || '0', 10)));
    if (p.action === 'admin') { requireAdmin_(p.admin); return json_(adminPayload_(p.scope === 'all')); }
    return json_({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents || '{}');
    if (b.action === 'practice') return json_(savePractice_(requireStudent_(b.k), b.date, b.done === true));
    if (b.action === 'attend') { requireAdmin_(b.admin); return json_(attend_(requireStudent_(b.k))); }
    if (b.action === 'record') { requireAdmin_(b.admin); return json_(saveRecord_(requireStudent_(b.k), b)); }
    return json_({ ok: false, error: '알 수 없는 요청' });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

/* =============================================================
 * 학생 화면 데이터
 * =========================================================== */
function studentPayload_(st) {
  var ss = SpreadsheetApp.openByUrl(st.sheetUrl);
  var rows = readLog_(ss);
  var sessions = rows.filter(function (r) { return r.isSession; });
  var latest = sessions.length ? sessions[sessions.length - 1] : null;
  var practice = readPractice_(ss);

  return {
    ok: true,
    name: st.name,
    today: todayStr_(),
    latest: latest ? lessonView_(latest) : null,
    materials: materialsView_(sessions),
    week: weekView_(practice),
    pass: passView_(sessions, rows.filter(function (r) { return r.isPlanned && r.row > (latest ? latest.row : 0); })),
    sessions: blockRows_(rows).reverse(),
    payments: paymentsView_(rows),
    repertoire: repertoire_(sessions),
    stats: statsView_(sessions)
  };
}

// 레슨 기록: 4회씩 한 묶음. 결제(D열)가 적힌 회차에서 새 묶음이 시작되고 '납부 완료',
// 결제 없이 4회가 지나 시작된 묶음은 '납부 전'으로 표시
function blockRows_(rows) {
  var block = -1, pos = 0, paid = false, out = [];
  rows.forEach(function (r) {
    if (r.isBreak) { out.push({ isBreak: true, content: '휴강' }); return; }
    if (!r.isSession) return;
    var item = { no: r.no, dateLabel: dateLabel_(r.date), content: [r.song, r.tech].filter(Boolean).join(' · ') };
    if (r.no === 0) {
      item.trial = true;                     // 0회차 = 시범 수업
    } else {
      if (r.pay) { block++; pos = 0; paid = true; item.pay = payLabel_(r); }
      else if (block < 0 || pos >= PASS_SIZE) { block++; pos = 0; paid = false; }
      pos++;
      item.block = block; item.pos = pos; item.paid = paid;
    }
    out.push(item);
  });
  return out;
}

function payLabel_(r) {
  return {
    date: Utilities.formatDate(r.pay, TZ, 'yyyy.MM.dd'),
    fee: r.fee ? Utilities.formatString('%s원', String(r.fee).replace(/\B(?=(\d{3})+(?!\d))/g, ',')) : ''
  };
}

// 수강료 납부 내역 (최근순) — 레슨 기록의 묶음 색과 같은 번호(block)를 함께 보냄
function paymentsView_(rows) {
  return blockRows_(rows).filter(function (r) { return r.pay; })
    .map(function (r) {
      return { date: r.pay.date, fee: r.pay.fee, from: r.no, to: r.no + PASS_SIZE - 1, block: r.block };
    }).reverse();
}

function lessonView_(r) {
  return { no: r.no, date: r.date, dateLabel: dateLabel_(r.date), song: r.song, tech: r.tech, task: r.task };
}

// 자료: 각 칸에서 가장 최근에 값이 있는 회차 기준
function materialsView_(sessions) {
  function latestWith(field) {
    for (var i = sessions.length - 1; i >= 0; i--) if (sessions[i][field]) return sessions[i];
    return null;
  }
  var sRow = latestWith('sheets'), vRow = latestWith('videos'), aRow = latestWith('audio');
  var audio = aRow ? parseLinks_(aRow.audio, '연습 음원') : [];
  return {
    sheets: sRow ? parseLinks_(sRow.sheets, (sRow.song || '') + ' 악보') : [],
    videos: vRow ? parseLinks_(vRow.videos, (vRow.song || '') + ' 영상') : [],
    audio: audio.map(function (a, i) { return { index: i, label: a.label }; }),
    audioSong: aRow ? (aRow.song || '') : '',
    bpm: aRow && aRow.bpm ? Number(aRow.bpm) : null,
    quick: aRow && aRow.quick ? String(aRow.quick).split(/[,\s]+/).map(Number).filter(function (n) { return n > 0; }) : []
  };
}

function weekView_(practice) {
  var now = new Date();
  var u = parseInt(Utilities.formatDate(now, TZ, 'u'), 10); // 1=월 … 7=일
  var today = todayStr_();
  var days = [];
  for (var i = 0; i < 7; i++) {
    var d = new Date(now.getTime() + (i - (u - 1)) * 86400000);
    var ds = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    days.push({
      date: ds,
      label: ['월', '화', '수', '목', '금', '토', '일'][i],
      day: Utilities.formatDate(d, TZ, 'd'),
      done: practice[ds] === true,
      future: ds > today
    });
  }
  return { days: days, count: days.filter(function (d) { return d.done; }).length };
}

// 결제일(D)이 적힌 마지막 회차부터 4회 단위
function passView_(sessions, planned) {
  // 0회차는 시범 수업 → 수업료 회차 계산에서 제외
  sessions = sessions.filter(function (r) { return r.no !== 0; });
  if (!sessions.length) return null;
  var prepaid = (planned || []).some(function (r) { return !!r.pay; });
  var startIdx = -1;
  for (var i = sessions.length - 1; i >= 0; i--) if (sessions[i].pay) { startIdx = i; break; }
  if (startIdx < 0) return null;
  var start = sessions[startIdx];
  var count = sessions.length - startIdx;
  return {
    done: Math.min(count, PASS_SIZE),
    total: PASS_SIZE,
    from: start.no,
    to: start.no + PASS_SIZE - 1,
    paidLabel: Utilities.formatDate(start.pay, TZ, 'M/d'),
    payDue: count >= PASS_SIZE && !prepaid
  };
}

function statsView_(sessions) {
  if (!sessions.length) return { total: 0, since: '', years: 0 };
  var regular = sessions.filter(function (r) { return r.no !== 0; }); // 시범 수업(0회차) 제외
  var first = sessions[0].date;
  var years = first ? Math.floor((Date.now() - first.getTime()) / (365.25 * 86400000)) : 0;
  return {
    total: regular.length,
    since: first ? Utilities.formatDate(first, TZ, 'yyyy.MM.dd') : '',
    years: years
  };
}

// 레퍼토리: E열(곡)을 '/'·'·' 로 나눠 곡 이름만 남기고 묶음 (자동 추정)
function repertoire_(sessions) {
  var groups = {}, order = [];
  sessions.forEach(function (r) {
    if (!r.song) return;
    String(r.song).split(/[\/·+]|\.\s/).forEach(function (part) {
      var title = normalizeSong_(part);
      if (!title) return;
      var key = title.toLowerCase().replace(/\s+/g, '');
      if (!groups[key]) { groups[key] = { title: title, first: r, last: r, count: 1 }; order.push(key); }
      else if (groups[key].last.no !== r.no) { groups[key].last = r; groups[key].count++; }
    });
  });
  return order.map(function (k) { return groups[k]; })
    .sort(function (a, b) { return b.last.no - a.last.no; })
    .slice(0, 60)
    .map(function (g) {
      return {
        when: g.last.date ? Utilities.formatDate(g.last.date, TZ, 'yy.MM') : '',
        title: g.title,
        meta: g.count === 1 ? g.last.no + '회차' : g.count + '회 레슨 · 마지막 ' + g.last.no + '회차'
      };
    });
}

// 곡 이름만 남기기 (자동 추정이라 일부는 정확하지 않을 수 있음)
function normalizeSong_(s) {
  var t = String(s).replace(/^[-\s]+/, '').replace(/\(.*?\)/g, '').trim();
  t = t.replace(/,.*$/, '');                 // 쉼표 뒤 설명 제거
  t = t.replace(/(\S)-\s*.*$/, '$1');        // "곡이름-설명" (하이픈 앞에 공백 없음) → 설명 제거. "가수 - 곡" 은 유지
  var suffix = /\s*(복습|총정리|완곡.*|완료|전체|코러스.*|인트로.*|브릿지.*|아웃트로.*|마지막.*|앞\s*부분|뒷\s*부분|\S*부분.*|.*부터.*|.*까지.*|쉬운 버전|풀 버전|\d+마디.*|\d+페이지.*)\s*$/;
  for (var i = 0; i < 3; i++) t = t.replace(suffix, '').replace(/[\s\-–:,.]+$/, '').trim();
  t = t.replace(/([A-Za-z가-힣])\s?\d$/, '$1').trim(); // "Life of a Fool 2" → "Life of a Fool"
  if (!t || t.length < 2) return '';
  if (/^(과제|개관|휴강|리듬연습|곡써온거)/.test(t)) return '';
  return t;
}

/* =============================================================
 * 연습 음원
 * =========================================================== */
function audioPayload_(st, index) {
  var ss = SpreadsheetApp.openByUrl(st.sheetUrl);
  var sessions = readLog_(ss).filter(function (r) { return r.isSession; });
  var row = null;
  for (var i = sessions.length - 1; i >= 0; i--) if (sessions[i].audio) { row = sessions[i]; break; }
  if (!row) throw new Error('등록된 음원이 없습니다');
  var links = parseLinks_(row.audio, '');
  var link = links[index];
  if (!link) throw new Error('음원을 찾을 수 없습니다');
  var id = driveId_(link.url);
  if (!id) throw new Error('드라이브 파일 주소가 아닙니다');
  var file = DriveApp.getFileById(id);
  if (file.getSize() > MAX_AUDIO_BYTES) throw new Error('파일이 너무 큽니다 (20MB 이하 MP3 권장)');
  var blob = file.getBlob();
  return { ok: true, mime: blob.getContentType() || 'audio/mpeg', data: Utilities.base64Encode(blob.getBytes()) };
}

/* =============================================================
 * 쓰기: 연습 체크 / 출석 / 레슨 기록
 * =========================================================== */
function savePractice_(st, date, done) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('날짜 형식 오류');
  var week = weekView_({}).days.map(function (d) { return d.date; });
  if (week.indexOf(date) < 0) throw new Error('이번 주 날짜만 체크할 수 있습니다');
  if (date > todayStr_()) throw new Error('오늘 이후 날짜는 체크할 수 없습니다');

  return withLock_(function () {
    var ss = SpreadsheetApp.openByUrl(st.sheetUrl);
    var sh = ss.getSheetByName(PRACTICE_TAB);
    if (!sh) {
      sh = ss.insertSheet(PRACTICE_TAB);
      sh.getRange(1, 1, 1, 2).setValues([['날짜', '연습함']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    var n = sh.getLastRow() - 1;
    var vals = n > 0 ? sh.getRange(2, 1, n, 1).getValues() : [];
    for (var i = 0; i < vals.length; i++) {
      if (cellDateStr_(vals[i][0]) === date) {
        sh.getRange(i + 2, 2).setValue(done);
        return { ok: true };
      }
    }
    sh.appendRow([noon_(date), done]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy-mm-dd');
    return { ok: true };
  });
}

function attend_(st) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openByUrl(st.sheetUrl);
    var sh = logSheet_(ss);
    var rows = readLog_(ss);
    var sessions = rows.filter(function (r) { return r.isSession; });
    var latest = sessions[sessions.length - 1];
    var today = todayStr_();
    var todayRow = rows.filter(function (r) {
      return r.date && Utilities.formatDate(r.date, TZ, 'yyyy-MM-dd') === today;
    })[0];
    var no = latest ? latest.no + 1 : 0;

    if (todayRow) {
      // 이미 오늘 행이 있음 (미리 적어둔 예정 행 포함)
      var cur = sh.getRange(todayRow.row, C.DONE).getValue();
      if (cur === true && todayRow.no !== null) return { ok: true, already: true, no: todayRow.no };
      sh.getRange(todayRow.row, C.DONE).setValue(true);
      if (todayRow.no === null || (latest && todayRow.row !== latest.row && todayRow.no <= latest.no)) {
        sh.getRange(todayRow.row, C.NO).setValue(todayRow.no === null ? no : todayRow.no);
      }
      return { ok: true, no: todayRow.no === null ? no : todayRow.no };
    }

    var r = latest ? latest.row + 1 : 2;
    sh.insertRowBefore(r);
    sh.getRange(r, C.DONE).setValue(true);
    sh.getRange(r, C.NO).setValue(no);
    sh.getRange(r, C.DATE).setValue(noon_(today)).setNumberFormat('yyyy-mm-dd');
    sh.getRange(r, C.PAY).setNumberFormat('yyyy-mm-dd');
    return { ok: true, no: no };
  });
}

function saveRecord_(st, b) {
  return withLock_(function () {
    var ss = SpreadsheetApp.openByUrl(st.sheetUrl);
    var sh = logSheet_(ss);
    var sessions = readLog_(ss).filter(function (r) { return r.isSession; });
    if (!sessions.length) throw new Error('회차가 없습니다. 먼저 출석을 눌러주세요');
    var row = sessions[sessions.length - 1].row;
    var set = function (col, v) { if (v !== undefined && v !== null) sh.getRange(row, col).setValue(v); };
    set(C.SONG, clean_(b.song));
    set(C.TECH, clean_(b.tech));
    set(C.TASK, clean_(b.task));
    set(C.SHEETS, clean_(b.sheets));
    set(C.VIDEOS, clean_(b.videos));
    set(C.AUDIO, clean_(b.audio));
    set(C.BPM, b.bpm ? Number(b.bpm) : '');
    set(C.QUICK, clean_(b.quick));
    if (b.paid === true) {
      sh.getRange(row, C.PAY).setValue(noon_(todayStr_())).setNumberFormat('yyyy-mm-dd');
      var fee = Number(String(b.fee || '').replace(/[^0-9]/g, ''));
      if (fee > 0) sh.getRange(row, C.FEE).setValue(fee).setNumberFormat('#,##0');
    }
    return { ok: true, no: sessions[sessions.length - 1].no };
  });
}

/* =============================================================
 * 선생님 화면 데이터
 * =========================================================== */
function adminPayload_(all) {
  var todayDow = DOW[parseInt(Utilities.formatDate(new Date(), TZ, 'u'), 10) % 7];
  var students = listStudents_().filter(function (s) {
    return all || String(s.day).indexOf(todayDow) >= 0;
  });
  var today = todayStr_();
  var base = props_().getProperty('APP_URL') || '';

  var out = students.map(function (st) {
    try {
      var ss = SpreadsheetApp.openByUrl(st.sheetUrl);
      var rowsAll = readLog_(ss);
      var sessions = rowsAll.filter(function (r) { return r.isSession; });
      var latest = sessions[sessions.length - 1] || null;
      var planned = rowsAll.filter(function (r) { return r.isPlanned && r.row > (latest ? latest.row : 0); });
      var attended = !!(latest && latest.date && Utilities.formatDate(latest.date, TZ, 'yyyy-MM-dd') === today);
      var practice = readPractice_(ss);
      var last7 = 0;
      for (var i = 1; i <= 7; i++) {
        var ds = Utilities.formatDate(new Date(Date.now() - i * 86400000), TZ, 'yyyy-MM-dd');
        if (practice[ds] === true) last7++;
      }
      return {
        key: st.key, name: st.name, time: st.time, day: st.day,
        next: latest ? (attended ? latest.no : latest.no + 1) : 0,
        attended: attended,
        practice: last7,
        pass: passView_(sessions, planned),
        link: base ? base + '?k=' + st.key : '',
        latest: latest ? {
          no: latest.no, song: latest.song, tech: latest.tech, task: latest.task,
          sheets: latest.sheets, videos: latest.videos, audio: latest.audio,
          bpm: latest.bpm || '', quick: latest.quick || ''
        } : null
      };
    } catch (err) {
      return { key: st.key, name: st.name, time: st.time, error: err.message };
    }
  });
  out.sort(function (a, b) { return String(a.time).localeCompare(String(b.time)); });
  return {
    ok: true,
    todayLabel: Utilities.formatDate(new Date(), TZ, 'M월 d일') + ' (' + todayDow + ')',
    scope: all ? 'all' : 'today',
    students: out
  };
}

/* =============================================================
 * 시트 읽기
 * =========================================================== */
function readLog_(ss) {
  var sh = logSheet_(ss);
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var vals = sh.getRange(2, 1, n, 15).getValues();
  var today = todayStr_();
  var out = [];
  vals.forEach(function (v, i) {
    var no = v[C.NO - 1];
    var hasNo = typeof no === 'number' || (/^\d+$/.test(String(no).trim()) && String(no).trim() !== '');
    var date = v[C.DATE - 1] instanceof Date ? v[C.DATE - 1] : null;
    // 레슨 = 회차 번호 + 날짜가 있고 오늘 이전. 미리 적어둔 다음 회차(미래 날짜·날짜 없음)는 예정 행
    var isSession = hasNo && !!date && Utilities.formatDate(date, TZ, 'yyyy-MM-dd') <= today;
    var isBreak = !hasNo && String(v[C.SONG - 1]).trim() === '휴강';
    var isPlanned = !isSession && !isBreak && (hasNo || !!date || String(v[C.SONG - 1]).trim() !== '');
    if (!isSession && !isBreak && !isPlanned) return;
    out.push({
      row: i + 2,
      isSession: isSession, isBreak: isBreak, isPlanned: isPlanned,
      no: hasNo ? Number(no) : null,
      date: date,
      pay: v[C.PAY - 1] instanceof Date ? v[C.PAY - 1] : null,
      song: str_(v[C.SONG - 1]), tech: str_(v[C.TECH - 1]), task: str_(v[C.TASK - 1]),
      sheets: str_(v[C.SHEETS - 1]), videos: str_(v[C.VIDEOS - 1]), audio: str_(v[C.AUDIO - 1]),
      bpm: v[C.BPM - 1], quick: str_(v[C.QUICK - 1]),
      fee: typeof v[C.FEE - 1] === 'number' ? v[C.FEE - 1] : null
    });
  });
  return out;
}

function readPractice_(ss) {
  var sh = ss.getSheetByName(PRACTICE_TAB);
  var map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var ds = cellDateStr_(r[0]);
    if (ds) map[ds] = r[1] === true || String(r[1]).toUpperCase() === 'TRUE';
  });
  return map;
}

function logSheet_(ss) {
  return ss.getSheetByName(LOG_TAB) || ss.getSheets()[0];
}

/* =============================================================
 * 학생/관리자 확인
 * =========================================================== */
function listStudents_() {
  var sh = studentSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 7).getValues()
    .filter(function (r) { return r[0] === true && r[1] && r[2] && r[5]; })
    .map(function (r) {
      return { name: String(r[1]), sheetUrl: toSheetUrl_(String(r[2])), day: String(r[3]),
               time: r[4] instanceof Date ? Utilities.formatDate(r[4], TZ, 'HH:mm') : String(r[4]),
               key: String(r[5]) };
    });
}

function requireStudent_(key) {
  if (!key || String(key).length < 12) throw new Error('등록된 링크가 아닙니다');
  var st = listStudents_().filter(function (s) { return s.key === String(key); })[0];
  if (!st) throw new Error('등록된 링크가 아닙니다');
  return st;
}

function requireAdmin_(key) {
  var real = props_().getProperty('ADMIN_KEY');
  if (!real || key !== real) throw new Error('선생님 링크가 아닙니다');
}

function ensureAdminKey_() {
  if (!props_().getProperty('ADMIN_KEY')) props_().setProperty('ADMIN_KEY', randomKey_() + randomKey_());
}

function refreshLinkColumn_() {
  var sh = studentSheet_();
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var base = props_().getProperty('APP_URL') || '';
  var keys = sh.getRange(2, 6, n, 1).getValues();
  sh.getRange(2, 7, n, 1).setValues(keys.map(function (k) {
    return [k[0] && base ? base + '?k=' + k[0] : ''];
  }));
}

function studentSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENT_TAB);
  if (!sh) throw new Error('[학생] 탭이 없습니다. "B4 앱 > 학생 탭 만들기"를 먼저 실행하세요');
  return sh;
}

/* =============================================================
 * 유틸
 * =========================================================== */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function props_() { return PropertiesService.getScriptProperties(); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function noon_(ds) { return new Date(ds + 'T12:00:00+09:00'); }
function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function clean_(v) { return v === undefined ? undefined : String(v || '').trim(); }

function dateLabel_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, TZ, 'yyyy.MM.dd') + ' (' + DOW[parseInt(Utilities.formatDate(d, TZ, 'u'), 10) % 7] + ')';
}

function cellDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

// 한 칸에 여러 링크: 줄바꿈으로 구분. "제목 | 주소" 형식이면 제목 사용
function parseLinks_(text, fallback) {
  return String(text || '').split(/\n+/).map(function (line, i, arr) {
    line = line.trim();
    if (!line) return null;
    var parts = line.split('|');
    var url = parts.length > 1 ? parts.slice(1).join('|').trim() : line;
    var label = parts.length > 1 ? parts[0].trim() : (fallback ? fallback + (arr.length > 1 ? ' ' + (i + 1) : '') : '');
    return /^https?:\/\//.test(url) ? { label: label || url, url: url } : null;
  }).filter(Boolean);
}

function driveId_(url) {
  var m = String(url).match(/\/d\/([a-zA-Z0-9_-]{20,})/) || String(url).match(/[?&]id=([a-zA-Z0-9_-]{20,})/);
  return m ? m[1] : null;
}

function toSheetUrl_(s) {
  if (/^https:\/\//.test(s)) return s;
  return 'https://docs.google.com/spreadsheets/d/' + s + '/edit';
}

function randomKey_() {
  var chars = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var out = '';
  for (var i = 0; i < 16; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}
