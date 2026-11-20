/**
 * 반찬가게 데이터 서버 (Google Apps Script)
 * - 손님 사이트:  웹앱주소           → 오늘 파는 반찬 목록(JSON)
 * - 엄마 입력판:  웹앱주소?page=mom  → 숫자만 누르면 자동 저장
 * - 관리용:       웹앱주소?page=admin → 판매 여부·수량·가격 수정
 * - 계산대 앱:    웹앱주소로 POST     → 라벨 인쇄(재고+)·계산(판매 기록)·폐기·할인
 * 구글 시트 "반찬가게 입력판"의 [확장 프로그램 → Apps Script]에 붙여 넣어 사용합니다.
 */
const MENU_SHEET = '반찬목록';
const SHOP_SHEET = '가게정보';
const PACK_SHEET = '라벨팩';
const SALE_SHEET = '판매기록';
const TZ = 'Asia/Seoul';
// 반찬목록 열 순서: 오늘판매, 분류, 반찬이름, 용량, 가격, 남은수량, 만든때, 사진, 소비기한(일), 원재료, 알레르기, 식품유형
const COL = { on: 0, cat: 1, name: 2, unit: 3, price: 4, left: 5, made: 6, photo: 7, days: 8, ingr: 9, allergy: 10, type: 11 };
const MENU_EXTRA_HEAD = ['소비기한(일)', '원재료', '알레르기', '식품유형'];
const MENU_COLS = COL.type + 1;
const PACK_HEAD = ['코드', '반찬이름', '용량', '가격', '만든날', '소비기한', '상태', '처리시각', '할인가'];
const SALE_HEAD = ['영수번호', '시각', '반찬이름', '가격', '결제', '코드', '상태'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.page === 'mom') {
    return HtmlService.createHtmlOutputFromFile('mom')
      .setTitle('마니반찬 오늘 반찬')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }
  if (p.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('반찬 입력판')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }
  return json_(publicData_());
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function book_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function shop_() {
  const rows = book_().getSheetByName(SHOP_SHEET).getDataRange().getValues();
  const map = {};
  rows.slice(1).forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1] ?? '').trim(); });
  return map;
}

function items_() {
  const sh = book_().getSheetByName(MENU_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const nc = Math.min(MENU_COLS, sh.getMaxColumns());
  return sh.getRange(2, 1, last - 1, nc).getValues().map((r, i) => {
    const d = r[COL.days];
    return {
      row: i + 2,
      on: r[COL.on] === true || String(r[COL.on]).toUpperCase() === 'TRUE',
      cat: String(r[COL.cat] || '기타').trim(),
      name: String(r[COL.name] || '').trim(),
      unit: String(r[COL.unit] || '').trim(),
      price: Number(r[COL.price]) || 0,
      left: Math.max(0, Math.round(Number(r[COL.left]) || 0)),
      made: String(r[COL.made] || '').trim(),
      photo: String(r[COL.photo] || '').trim(),
      days: d === '' || d === undefined || isNaN(Number(d)) ? 3 : Math.max(0, Math.min(60, Math.round(Number(d)))),
      ingr: String(r[COL.ingr] || '').trim(),
      allergy: String(r[COL.allergy] || '').trim(),
      type: String(r[COL.type] || '').trim()
    };
  }).filter(it => it.name);
}

function publicData_() {
  const s = shop_(), all = items_();
  return {
    shop: { name: s['가게이름'] || '', tagline: s['한줄소개'] || '', hours: s['영업시간'] || '', phone: s['전화번호'] || '', chat: s['오픈채팅'] || '', address: s['주소'] || '' },
    updatedAt: PropertiesService.getScriptProperties().getProperty('updatedAt') || '',
    items: all.filter(it => it.on).map(it => ({ cat: it.cat, name: it.name, unit: it.unit, price: it.price, left: it.left, made: it.made, photo: it.photo })),
    // 매장 게시 안내판(/notice/)용 표시사항: 오늘 안 파는 반찬도 포함
    info: all.map(it => ({ on: it.on, cat: it.cat, name: it.name, unit: it.unit, type: it.type, ingr: it.ingr, allergy: it.allergy, days: it.days }))
  };
}

/* ---------- 엄마 입력판에서 부르는 함수 ---------- */
function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) throw new Error('비밀번호를 너무 많이 틀렸어요. 10분 뒤에 다시 해 주세요.');
  if (String(pin || '') !== String(shop_()['엄마비밀번호'] || '')) {
    cache.put('pinFails', String(fails + 1), 600);
    throw new Error('비밀번호가 맞지 않아요.');
  }
}

function touch_() {
  const now = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty('updatedAt', now);
  return now;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function adminLoad(pin) {
  checkPin_(pin);
  return { shopName: shop_()['가게이름'] || '', items: items_(), updatedAt: PropertiesService.getScriptProperties().getProperty('updatedAt') || '' };
}

/** changes: [{row, on, left, price}] */
function adminSave(pin, changes) {
  checkPin_(pin);
  return withLock_(() => {
    const sh = book_().getSheetByName(MENU_SHEET);
    const last = sh.getLastRow();
    (changes || []).forEach(c => {
      const row = Math.round(Number(c.row));
      if (!(row >= 2 && row <= last)) return;
      sh.getRange(row, COL.on + 1).setValue(!!c.on);
      sh.getRange(row, COL.price + 1).setValue(Math.max(0, Math.round(Number(c.price) || 0)));
      sh.getRange(row, COL.left + 1).setValue(Math.max(0, Math.round(Number(c.left) || 0)));
    });
    return { updatedAt: touch_() };
  });
}

/* ---------- 계산대 앱 (POST, 본문은 JSON 글자) ---------- */
const POS_ACTIONS = {
  load: () => ({}),
  label: req => posLabel_(req),
  sell: req => posSell_(req),
  cancel: req => posCancel_(req),
  discard: req => posDiscard_(req),
  discount: req => posDiscount_(req)
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkPin_(req.pin);
    const fn = POS_ACTIONS[req.action];
    if (!fn) throw new Error('알 수 없는 요청이에요.');
    ensurePos_();
    const result = fn(req) || {};
    out = Object.assign({ ok: true }, result, posLoad_());
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return json_(out);
}

function fmt_(d, p) { return Utilities.formatDate(d, TZ, p); }
function nowStr_() { return fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return fmt_(new Date(), 'yyyy-MM-dd'); }
function addDays_(ymd, n) {
  const p = String(ymd).split('-').map(Number);
  return Utilities.formatDate(new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)), 'UTC', 'yyyy-MM-dd');
}
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
/** 시트가 글자를 날짜로 바꿔 버렸어도 'yyyy-MM-dd' / 'yyyy-MM-dd HH:mm:ss' 글자로 되돌림 */
function cellStr_(v, withTime) {
  if (isDate_(v)) return Utilities.formatDate(v, book_().getSpreadsheetTimeZone(), withTime ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  return String(v ?? '').trim();
}

function makeSheet_(name, head, textCols) {
  const b = book_();
  let sh = b.getSheetByName(name);
  if (!sh) {
    sh = b.insertSheet(name);
    textCols.forEach(c => sh.getRange(c).setNumberFormat('@'));
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 처음 한 번: 라벨팩·판매기록 탭을 만들고 반찬목록에 소비기한·원재료·알레르기 칸 제목을 붙임 */
function ensurePos_() {
  makeSheet_(PACK_SHEET, PACK_HEAD, ['A:A', 'E:F', 'H:H']);
  makeSheet_(SALE_SHEET, SALE_HEAD, ['A:B', 'F:F']);
  const m = book_().getSheetByName(MENU_SHEET);
  if (m.getMaxColumns() < MENU_COLS) m.insertColumnsAfter(m.getMaxColumns(), MENU_COLS - m.getMaxColumns());
  const n = MENU_EXTRA_HEAD.length, head = m.getRange(1, COL.days + 1, 1, n).getValues()[0];
  if (head.some(h => !String(h).trim())) {
    m.getRange(1, COL.days + 1, 1, n).setValues([head.map((h, i) => String(h).trim() || MENU_EXTRA_HEAD[i])]);
  }
}

function rows_(name, width) {
  const sh = book_().getSheetByName(name);
  const last = sh.getLastRow();
  return { sh, data: last < 2 ? [] : sh.getRange(2, 1, last - 1, width).getValues() };
}

function packObj_(r) {
  return { code: String(r[0]), name: String(r[1]), unit: String(r[2]), price: Number(r[3]) || 0, made: cellStr_(r[4]), exp: cellStr_(r[5]), st: String(r[6]), at: cellStr_(r[7], true), sale: Number(r[8]) || 0 };
}
function saleObj_(r) {
  return { rid: String(r[0]), t: cellStr_(r[1], true), name: String(r[2]), price: Number(r[3]) || 0, pay: String(r[4]), code: String(r[5]), st: String(r[6]) };
}

/** 모든 응답에 붙는 현재 상태: 반찬, 재고 팩, 최근 7일 판매·폐기 */
function posLoad_() {
  const today = today_(), from = addDays_(today, -6), s = shop_();
  const packs = rows_(PACK_SHEET, PACK_HEAD.length).data.map(packObj_);
  return {
    today,
    now: nowStr_(),
    shop: { name: s['가게이름'] || '', phone: s['전화번호'] || '', discount: Number(s['마감할인율']) || 30 },
    items: items_(),
    packs: packs.filter(p => p.st === '재고'),
    gone: packs.filter(p => p.st === '폐기' && p.at.slice(0, 10) >= from).map(p => ({ code: p.code, name: p.name, at: p.at })),
    sales: rows_(SALE_SHEET, SALE_HEAD.length).data.map(saleObj_).filter(x => x.st === '정상' && x.t.slice(0, 10) >= from)
  };
}

function nextCodes_(n) {
  const props = PropertiesService.getScriptProperties();
  const day = fmt_(new Date(), 'yyMMdd'), key = 'seq' + day;
  props.getKeys().filter(k => /^seq\d{6}$/.test(k) && k !== key).forEach(k => props.deleteProperty(k));
  let s = Number(props.getProperty(key) || 0);
  const out = [];
  for (let i = 0; i < n; i++) { s++; out.push('M' + day + '-' + String(s).padStart(3, '0')); }
  props.setProperty(key, String(s));
  return out;
}

/** 남은수량 더하기/빼기. turnOn이면 늘어난 반찬을 '오늘판매'로 켬 */
function addLeft_(turnOn, delta) {
  const m = book_().getSheetByName(MENU_SHEET);
  const items = items_();
  Object.keys(delta).forEach(name => {
    const it = items.find(x => x.name === name);
    if (!it) return;
    m.getRange(it.row, COL.left + 1).setValue(Math.max(0, it.left + delta[name]));
    if (delta[name] > 0 && turnOn) m.getRange(it.row, COL.on + 1).setValue(true);
  });
}

/** 라벨 인쇄 = 팩 등록(재고 +qty). req: {row, name, qty} */
function posLabel_(req) {
  const qty = Math.round(Number(req.qty));
  if (!(qty >= 1 && qty <= 60)) throw new Error('라벨은 한 번에 1~60장까지예요.');
  return withLock_(() => {
    const it = items_().find(x => x.row === Number(req.row) && x.name === String(req.name));
    if (!it) throw new Error('반찬을 찾지 못했어요. 새로고침한 뒤 다시 해 주세요.');
    const made = today_(), exp = addDays_(made, it.days), t = nowStr_();
    const rows = nextCodes_(qty).map(c => [c, it.name, it.unit, it.price, made, exp, '재고', t, '']);
    const sh = book_().getSheetByName(PACK_SHEET);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, PACK_HEAD.length).setValues(rows);
    addLeft_(true, { [it.name]: qty });
    touch_();
    return { printed: rows.map(packObj_).map(p => Object.assign(p, { ingr: it.ingr, allergy: it.allergy })) };
  });
}

/** 계산. req: {pay:'카드'|'현금', lines:[{code} 또는 {name}]} */
function posSell_(req) {
  const pay = req.pay === '현금' ? '현금' : '카드';
  const lines = Array.isArray(req.lines) ? req.lines.slice(0, 100) : [];
  if (!lines.length) throw new Error('담은 반찬이 없어요.');
  return withLock_(() => {
    const { sh: ps, data } = rows_(PACK_SHEET, PACK_HEAD.length);
    const idx = {};
    data.forEach((r, i) => { idx[String(r[0])] = i; });
    const items = items_();
    const t = nowStr_(), rid = 'R' + fmt_(new Date(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    const sales = [], warn = [], dec = {};
    lines.forEach(l => {
      const code = String(l.code || '').trim();
      if (code) {
        const i = idx[code];
        if (i === undefined) { warn.push(code + ' 라벨은 기록에 없어요'); return; }
        const r = data[i];
        if (String(r[6]) !== '재고') { warn.push(r[1] + ' ' + code + '는 이미 ' + r[6] + '로 기록돼 있어요'); return; }
        r[6] = '판매';
        ps.getRange(i + 2, 7, 1, 2).setValues([['판매', t]]);
        sales.push([rid, t, String(r[1]), Number(r[8]) || Number(r[3]) || 0, pay, code, '정상']);
        dec[r[1]] = (dec[r[1]] || 0) - 1;
      } else {
        const it = items.find(x => x.name === String(l.name || ''));
        if (!it) { warn.push((l.name || '이름 없는 반찬') + '은 목록에 없어요'); return; }
        sales.push([rid, t, it.name, it.price, pay, '', '정상']);
        dec[it.name] = (dec[it.name] || 0) - 1;
      }
    });
    if (!sales.length) throw new Error(warn.join(' / ') || '기록할 반찬이 없어요.');
    const ss = book_().getSheetByName(SALE_SHEET);
    ss.getRange(ss.getLastRow() + 1, 1, sales.length, SALE_HEAD.length).setValues(sales);
    addLeft_(false, dec);
    touch_();
    return { receipt: { rid, pay, count: sales.length, sum: sales.reduce((a, s) => a + s[3], 0) }, warn };
  });
}

/** 오늘 계산 취소 (실수했을 때). req: {rid} */
function posCancel_(req) {
  const rid = String(req.rid || '');
  return withLock_(() => {
    const { sh: ss, data } = rows_(SALE_SHEET, SALE_HEAD.length);
    const today = today_(), inc = {}, codes = [];
    let n = 0;
    data.forEach((r, i) => {
      if (String(r[0]) !== rid || String(r[6]) !== '정상') return;
      if (cellStr_(r[1], true).slice(0, 10) !== today) throw new Error('오늘 계산만 취소할 수 있어요.');
      ss.getRange(i + 2, 7).setValue('취소');
      inc[r[2]] = (inc[r[2]] || 0) + 1;
      if (String(r[5])) codes.push(String(r[5]));
      n++;
    });
    if (!n) throw new Error('취소할 계산을 찾지 못했어요.');
    if (codes.length) {
      const { sh: ps, data: pd } = rows_(PACK_SHEET, PACK_HEAD.length);
      pd.forEach((r, i) => { if (codes.indexOf(String(r[0])) >= 0 && String(r[6]) === '판매') ps.getRange(i + 2, 7, 1, 2).setValues([['재고', nowStr_()]]); });
    }
    addLeft_(false, inc);
    touch_();
    return { cancelled: n };
  });
}

function findPack_(code) {
  const { sh, data } = rows_(PACK_SHEET, PACK_HEAD.length);
  const i = data.findIndex(r => String(r[0]) === String(code || ''));
  if (i < 0) throw new Error('라벨 ' + code + '을 찾지 못했어요.');
  if (String(data[i][6]) !== '재고') throw new Error(data[i][1] + '는 이미 ' + data[i][6] + '로 기록돼 있어요.');
  return { sh, row: i + 2, r: data[i] };
}

/** 폐기. req: {code} */
function posDiscard_(req) {
  return withLock_(() => {
    const p = findPack_(req.code);
    p.sh.getRange(p.row, 7, 1, 2).setValues([['폐기', nowStr_()]]);
    addLeft_(false, { [p.r[1]]: -1 });
    touch_();
    return { discarded: String(req.code) };
  });
}

/** 마감 할인가 설정(0이면 해제). req: {code, price} */
function posDiscount_(req) {
  return withLock_(() => {
    const p = findPack_(req.code);
    const price = Math.max(0, Math.round(Number(req.price) || 0));
    p.sh.getRange(p.row, 9).setValue(price || '');
    return { discounted: String(req.code), price };
  });
}
