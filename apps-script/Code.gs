/**
 * 차곡차곡 가계부: 구글 시트 저장소
 *
 * 1) 이 코드를 구글 시트의 [확장 프로그램] → [Apps Script]에 붙여 넣으세요.
 * 2) [프로젝트 설정] → [스크립트 속성]에 PASSWORD 라는 이름으로 비밀번호를 추가하세요.
 *    (비밀번호는 이 코드나 GitHub에 적지 마세요.)
 * 3) setup 함수를 한 번 실행해 시트를 만드세요.
 * 4) [배포] → [새 배포] → 웹 앱, 실행: 나, 액세스: 모든 사용자 → 웹 앱 주소를 차곡차곡 설정에 넣으세요.
 */

const TABLES = {
  tx: {
    name: '내역',
    cols: ['id', '날짜', '시간', '금액', '가게·내용', '카테고리', '결제수단', '결제방식', '취소·환불', '메모', '출처', 'updatedAt', 'deleted', 'json'],
    row: r => [r.id, r.date, r.time || '', r.amount, r.merchant || '', r.category || '', r.method || '', r.pay || '', r.refund ? 'Y' : '', r.memo || '', r.source || '', r.updatedAt || 0, r.deleted ? 'Y' : '', JSON.stringify(r)]
  },
  savings: {
    name: '저축',
    cols: ['id', '날짜', '금액', '목표 id', '방법', '외화', '적용 환율', '메모', 'updatedAt', 'deleted', 'json'],
    row: r => [r.id, r.date, r.amount, r.goalId || '', r.method || '', r.cur ? r.cur + ' ' + r.fx : '', r.rate || '', r.memo || '', r.updatedAt || 0, r.deleted ? 'Y' : '', JSON.stringify(r)]
  },
  goals: {
    name: '저축 목표',
    cols: ['id', '이름', '목표 금액', 'updatedAt', 'deleted', 'json'],
    row: r => [r.id, r.name, r.target, r.updatedAt || 0, r.deleted ? 'Y' : '', JSON.stringify(r)]
  },
  special: {
    name: '특별 지출',
    cols: ['id', '이름', '금액', '날짜', '결제수단', '메모', 'updatedAt', 'deleted', 'json'],
    row: r => [r.id, r.name, r.amount, r.date, r.method || '', r.memo || '', r.updatedAt || 0, r.deleted ? 'Y' : '', JSON.stringify(r)]
  },
  bills: {
    name: '카드 결제 금액',
    cols: ['id', '카드', '결제일', '결제 금액', 'updatedAt', 'deleted', 'json'],
    row: r => [r.id, r.cardId, r.billDate, r.amount, r.updatedAt || 0, r.deleted ? 'Y' : '', JSON.stringify(r)]
  },
  fixed: {
    name: '정기 지출',
    cols: ['id', '이름', '금액', '매달', '결제수단', 'updatedAt', 'deleted', 'json'],
    row: r => [r.id, r.name, r.amount, r.day, r.method || '', r.updatedAt || 0, r.deleted ? 'Y' : '', JSON.stringify(r)]
  }
};
const INBOX = { name: '문자함', cols: ['id', '받은 시각', '문자', '처리됨'] };
const SETTINGS = { name: '설정', cols: ['updatedAt', 'json'] };

function setup() {
  Object.keys(TABLES).forEach(k => sheet_(TABLES[k]));
  sheet_(INBOX);
  sheet_(SETTINGS);
  sheet_(IMAGES);
  const pw = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  Logger.log(pw ? '준비 완료. 이제 웹 앱으로 배포하세요.' : '스크립트 속성에 PASSWORD를 먼저 추가하세요.');
}

function doGet() {
  return out_({ ok: true, app: 'chagok' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: '요청 형식이 잘못됐어요' });
  }
  const pw = PropertiesService.getScriptProperties().getProperty('PASSWORD');
  if (!pw) return out_({ ok: false, error: '스크립트 속성에 PASSWORD를 먼저 넣어 주세요' });
  if (String(body.pw || '') !== pw) return out_({ ok: false, error: '비밀번호가 맞지 않아요' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (body.action) {
      case 'ping':
        return out_({ ok: true });
      case 'sms':
        return out_(addSms_(body.text));
      case 'load':
        return out_({ ok: true, data: loadAll_(), inbox: pendingSms_() });
      case 'save':
        saveAll_(body.data || {});
        if (body.ack && body.ack.length) ackSms_(body.ack);
        return out_({ ok: true, savedAt: new Date().toISOString() });
      default:
        return out_({ ok: false, error: '알 수 없는 요청이에요' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 문자함 (아이폰 단축어가 보내는 결제 문자) ---------- */
function addSms_(text) {
  text = String(text || '').trim();
  if (!text) return { ok: false, error: '문자 내용이 비어 있어요' };
  const sh = sheet_(INBOX);
  const id = Utilities.getUuid().slice(0, 12);
  sh.appendRow([id, new Date(), text, '']);
  return { ok: true, id: id };
}
function pendingSms_() {
  const sh = sheet_(INBOX);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 4).getValues()
    .filter(r => r[0] && !r[3])
    .map(r => ({ id: String(r[0]), at: r[1], text: String(r[2]) }));
}
function ackSms_(ids) {
  const sh = sheet_(INBOX);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const set = {};
  ids.forEach(id => set[String(id)] = true);
  const rng = sh.getRange(2, 1, n, 4);
  const vals = rng.getValues();
  vals.forEach(r => { if (set[String(r[0])]) r[3] = 'Y'; });
  rng.setValues(vals);
}

/* ---------- 불러오기 / 저장 ---------- */
function loadAll_() {
  const data = {};
  Object.keys(TABLES).forEach(k => {
    const t = TABLES[k];
    const sh = sheet_(t);
    const n = sh.getLastRow() - 1;
    const jc = t.cols.length;
    data[k] = n < 1 ? [] : sh.getRange(2, jc, n, 1).getValues()
      .map(r => { try { return JSON.parse(r[0]); } catch (e) { return null; } })
      .filter(Boolean);
  });
  const st = sheet_(SETTINGS);
  if (st.getLastRow() >= 2) {
    try { data.settings = JSON.parse(st.getRange(2, 2).getValue()); } catch (e) {}
  }
  data.images = loadImages_();
  return data;
}
function saveAll_(d) {
  Object.keys(TABLES).forEach(k => {
    if (!Array.isArray(d[k])) return;
    const t = TABLES[k];
    const sh = sheet_(t);
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, t.cols.length).clearContent();
    const rows = d[k].filter(r => r && r.id)
      .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')))
      .map(t.row);
    if (rows.length) sh.getRange(2, 1, rows.length, t.cols.length).setValues(rows);
  });
  if (d.images) saveImages_(d.images);
  if (d.settings) {
    const st = sheet_(SETTINGS);
    st.getRange(2, 1, 1, 2).setValues([[d.settings.updatedAt || 0, JSON.stringify(d.settings)]]);
  }
}

/* ---------- 카드 사진 (한 칸에 다 안 들어가서 나눠 저장) ---------- */
const IMAGES = { name: '카드 사진', cols: ['cardId', 'updatedAt', 'part', 'data'] };
const PART = 45000;
function loadImages_() {
  const sh = sheet_(IMAGES);
  const n = sh.getLastRow() - 1;
  const out = {};
  if (n < 1) return out;
  const rows = sh.getRange(2, 1, n, 4).getValues();
  const parts = {};
  rows.forEach(r => { if (!r[0]) return; const id = String(r[0]); (parts[id] = parts[id] || { updatedAt: +r[1] || 0, list: [] }).list.push([+r[2], String(r[3])]); });
  Object.keys(parts).forEach(id => { const p = parts[id]; p.list.sort((a, b) => a[0] - b[0]); out[id] = { updatedAt: p.updatedAt, data: p.list.map(x => x[1]).join('') }; });
  return out;
}
function saveImages_(imgs) {
  const sh = sheet_(IMAGES);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
  const rows = [];
  Object.keys(imgs || {}).forEach(id => {
    const im = imgs[id];
    if (!im) return;
    if (!im.data) { rows.push([id, im.updatedAt || 0, 0, '']); return; }
    for (let i = 0, k = 0; i < im.data.length; i += PART, k++) rows.push([id, im.updatedAt || 0, k, im.data.slice(i, i + PART)]);
  });
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
}

/* ---------- 도구 ---------- */
function sheet_(t) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, 1, t.cols.length).setValues([t.cols]).setFontWeight('bold').setBackground('#FAF4E8');
    sh.setFrozenRows(1);
    const jc = t.cols.indexOf('json');
    if (jc >= 0) sh.hideColumns(jc + 1);
  }
  return sh;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
