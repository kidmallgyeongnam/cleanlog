/**
 * 클린로그 v2 — 서버 (Google Apps Script)
 * 화면은 별도 주소(깃허브 페이지)에 있고, 이 서버는 사진을 구글 드라이브에 정리해 저장합니다.
 * ▼ 접속 번호(4~6자리)와 이름. 실제 번호는 Apps Script 편집기 안에서만 적습니다(공개 저장소에는 적지 않음).
 */
var TEAM = {
  '여기에_번호': '관리자'
};
var ROOT = '해솔클린_청소현장기록';
var DATA = '_클린로그_데이터.json';
var CATS = { before: '01_청소전', during: '02_청소중', after: '03_청소후', etc: '04_기타작업' };
var APP_URL = 'https://kidmallgyeongnam.github.io/cleanlog/';

function doGet() {
  return HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width,initial-scale=1"><p style="font:18px sans-serif;padding:24px">클린로그 주소가 바뀌었습니다.<br><a href="' + APP_URL + '" target="_top">새 클린로그 열기</a></p>').setTitle('클린로그');
}

function doPost(e) {
  var out;
  try {
    var b = JSON.parse(e.postData.contents);
    var name = auth_(b.pin);
    var f = API[b.fn];
    if (!f) throw new Error('알 수 없는 요청입니다.');
    out = { ok: true, data: f(b.args || {}, name) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** 접속 번호 확인. 틀린 시도가 10분에 30번 넘으면 잠깐 막습니다. */
function auth_(pin) {
  var c = CacheService.getScriptCache(), n = Number(c.get('fail') || 0);
  if (n >= 30) throw new Error('접속 시도가 너무 많습니다. 10분 뒤 다시 해 주세요.');
  var name = TEAM[String(pin)];
  if (!name) { c.put('fail', String(n + 1), 600); throw new Error('접속 번호가 맞지 않습니다.'); }
  return name;
}
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\\/:*?"<>|\n\r\t]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max || 40); }
function child_(parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function root_() { return child_(DriveApp.getRootFolder(), ROOT); }
function data_() { var it = root_().getFilesByName(DATA); return it.hasNext() ? JSON.parse(it.next().getBlob().getDataAsString('UTF-8')) : { sites: [], sessions: {} }; }
function saveData_(d) { var r = root_(), it = r.getFilesByName(DATA), t = JSON.stringify(d); if (it.hasNext()) it.next().setContent(t); else r.createFile(DATA, t, 'application/json'); }
function putFile_(folder, name, text, mime) { var it = folder.getFilesByName(name); if (it.hasNext()) it.next().setContent(text); else folder.createFile(name, text, mime || 'text/plain'); }
function sessionFolder_(sid) {
  var c = CacheService.getScriptCache(), id = c.get('s_' + sid);
  if (!id) { var s = data_().sessions[sid]; if (!s) throw new Error('현장 작업을 찾을 수 없습니다. 업로드를 다시 시작해 주세요.'); id = s.folderId; c.put('s_' + sid, id, 21600); }
  return DriveApp.getFolderById(id);
}

var API = {
  login: function (a, name) { return { name: name }; },

  /** 등록된 현장 목록 (현장 시작 화면에서 고르는 용도) */
  bootstrap: function (a, name) { return { name: name, sites: data_().sites || [] }; },

  /** 좌표 → 지역명(예: 부산강서구). 못 찾으면 빈 값 — 작업자가 직접 입력합니다. */
  region: function (a) {
    try {
      var r = Maps.newGeocoder().setLanguage('ko').reverseGeocode(Number(a.lat), Number(a.lng));
      var comps = r.results && r.results[0] ? r.results[0].address_components : [], si = '', gu = '';
      comps.forEach(function (c) {
        if (c.types.indexOf('administrative_area_level_1') >= 0) si = c.long_name;
        if (!gu && (c.types.indexOf('sublocality_level_1') >= 0 || c.types.indexOf('locality') >= 0)) gu = c.long_name;
      });
      si = si.replace(/(특별자치시|특별자치도|광역시|특별시)$/, '');
      return { region: clean_(si + (gu === si ? '' : gu), 20).replace(/ /g, '') };
    } catch (err) { return { region: '', err: String(err && err.message || err).slice(0, 200) }; }
  },

  /** 현장 작업 폴더 만들기. 같은 작업번호(sid)로 다시 불러도 폴더는 하나만 생깁니다. */
  ensureSession: function (a, leader) {
    var sid = clean_(a.sid, 40); if (!sid) throw new Error('작업 번호가 없습니다.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(a.date)) throw new Error('작업일이 올바르지 않습니다.');
    var site = clean_(a.siteName, 30), region = clean_(a.region, 20).replace(/ /g, '');
    if (!site) throw new Error('현장명을 입력해 주세요.');
    var lock = LockService.getScriptLock(); lock.waitLock(25000);
    try {
      var d = data_(), s = d.sessions[sid], folder;
      if (s) folder = DriveApp.getFolderById(s.folderId);
      else {
        var month = child_(child_(root_(), a.date.slice(0, 4) + '년'), Number(a.date.slice(5, 7)) + '월');
        var base = [a.date, region, site].filter(String).join('_'), name = base, n = 1;
        while (month.getFoldersByName(name).hasNext()) { n++; name = base + '_' + n; }
        folder = month.createFolder(name);
        Object.keys(CATS).forEach(function (k) { folder.createFolder(CATS[k]); });
        s = d.sessions[sid] = { sid: sid, folderId: folder.getId(), name: name, siteName: site, region: region, date: a.date, leader: leader, status: '업로드중', startedAt: a.startedAt || '', lat: a.lat || null, lng: a.lng || null };
        if (!(d.sites || []).some(function (x) { return x.name === site; })) { d.sites = d.sites || []; d.sites.push({ name: site, region: region, lat: a.lat || null, lng: a.lng || null }); }
        saveData_(d);
      }
      CacheService.getScriptCache().put('s_' + sid, s.folderId, 21600);
      return { name: s.name, url: folder.getUrl() };
    } finally { lock.releaseLock(); }
  },

  /** 사진 1장 저장. 같은 이름(사진 고유값 포함)이 이미 있으면 다시 저장하지 않습니다. */
  uploadPhoto: function (a) {
    var sub = CATS[a.cat]; if (!sub) throw new Error('사진 구분이 올바르지 않습니다.');
    var name = clean_(a.name, 60); if (!name || !a.base64) throw new Error('사진 내용이 없습니다.');
    var target = child_(sessionFolder_(clean_(a.sid, 40)), sub);
    if (target.getFilesByName(name).hasNext()) return { dup: true };
    target.createFile(Utilities.newBlob(Utilities.base64Decode(a.base64), a.mime || 'image/jpeg', name));
    return { dup: false };
  },

  /** 현장 작업 종료: 사진 수 확인, 사진목록·현장정보·블로그용 자료 저장, 완료 표시 */
  finishSession: function (a, leader) {
    var sid = clean_(a.sid, 40), folder = sessionFolder_(sid), counts = {}, total = 0;
    Object.keys(CATS).forEach(function (k) { var it = child_(folder, CATS[k]).getFiles(), n = 0; while (it.hasNext()) { it.next(); n++; } counts[k] = n; total += n; });
    var memo = a.memo || {}, special = String(memo.special || '').slice(0, 1000), work = String(memo.work || '').slice(0, 1000), consent = clean_(memo.consent, 6) || '미확인';
    var lock = LockService.getScriptLock(); lock.waitLock(25000);
    var s;
    try {
      var d = data_(); s = d.sessions[sid]; if (!s) throw new Error('현장 작업을 찾을 수 없습니다.');
      s.status = '완료'; s.endedAt = a.endedAt || new Date().toISOString(); s.counts = counts; s.total = total; saveData_(d);
    } finally { lock.releaseLock(); }
    putFile_(folder, '사진목록.json', JSON.stringify(a.manifest || [], null, 1), 'application/json');
    putFile_(folder, '현장정보.json', JSON.stringify({ 작업번호: sid, 현장명: s.siteName, 지역: s.region, 작업일: s.date, 담당: s.leader, 시작: s.startedAt, 종료: s.endedAt, 사진수: counts, 특이사항: special, 작업내용: work, 사진사용동의: consent }, null, 2), 'application/json');
    // 블로그용 자료에는 지역·작업 내용만. 현장명(건물 이름)과 좌표는 넣지 않습니다.
    putFile_(folder, '블로그용_자료.md', ['# 블로그용 현장 자료', '', '- 작업일: ' + s.date, '- 지역: ' + (s.region || '미입력'), '- 블로그 사진 사용 동의: ' + consent,
      '- 사진: 전 ' + counts.before + '장 / 중 ' + counts.during + '장 / 후 ' + counts.after + '장 / 기타 ' + counts.etc + '장', '', '## 현장 특이사항', special || '(없음)', '', '## 작업 내용', work || '(없음)', ''].join('\n'), 'text/markdown');
    putFile_(folder, '업로드완료.txt', '업로드 완료: ' + s.endedAt + ' / 총 ' + total + '장');
    return { counts: counts, total: total, url: folder.getUrl(), name: s.name };
  }
};
