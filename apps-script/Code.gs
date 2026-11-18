/**
 * 반찬가게 데이터 서버 (Google Apps Script)
 * - 손님 사이트:  웹앱주소           → 오늘 파는 반찬 목록(JSON)
 * - 엄마 입력판:  웹앱주소?page=admin → 비밀번호 입력 후 판매 여부·수량·가격 수정
 * 구글 시트 "반찬가게 입력판"의 [확장 프로그램 → Apps Script]에 붙여 넣어 사용합니다.
 */
const MENU_SHEET = '반찬목록';
const SHOP_SHEET = '가게정보';
// 반찬목록 열 순서: 오늘판매, 분류, 반찬이름, 용량, 가격, 남은수량, 만든때, 사진
const COL = { on: 0, cat: 1, name: 2, unit: 3, price: 4, left: 5, made: 6, photo: 7 };

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
  return ContentService.createTextOutput(JSON.stringify(publicData_()))
    .setMimeType(ContentService.MimeType.JSON);
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
  return sh.getRange(2, 1, last - 1, 8).getValues().map((r, i) => ({
    row: i + 2,
    on: r[COL.on] === true || String(r[COL.on]).toUpperCase() === 'TRUE',
    cat: String(r[COL.cat] || '기타').trim(),
    name: String(r[COL.name] || '').trim(),
    unit: String(r[COL.unit] || '').trim(),
    price: Number(r[COL.price]) || 0,
    left: Math.max(0, Math.round(Number(r[COL.left]) || 0)),
    made: String(r[COL.made] || '').trim(),
    photo: String(r[COL.photo] || '').trim()
  })).filter(it => it.name);
}

function publicData_() {
  const s = shop_();
  return {
    shop: { name: s['가게이름'] || '', tagline: s['한줄소개'] || '', hours: s['영업시간'] || '', phone: s['전화번호'] || '', chat: s['오픈채팅'] || '' },
    updatedAt: PropertiesService.getScriptProperties().getProperty('updatedAt') || '',
    items: items_().filter(it => it.on).map(({ row, on, ...rest }) => rest)
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

function adminLoad(pin) {
  checkPin_(pin);
  return { shopName: shop_()['가게이름'] || '', items: items_(), updatedAt: PropertiesService.getScriptProperties().getProperty('updatedAt') || '' };
}

/** changes: [{row, on, left, price}] */
function adminSave(pin, changes) {
  checkPin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = book_().getSheetByName(MENU_SHEET);
    const last = sh.getLastRow();
    (changes || []).forEach(c => {
      const row = Math.round(Number(c.row));
      if (!(row >= 2 && row <= last)) return;
      sh.getRange(row, COL.on + 1).setValue(!!c.on);
      sh.getRange(row, COL.price + 1).setValue(Math.max(0, Math.round(Number(c.price) || 0)));
      sh.getRange(row, COL.left + 1).setValue(Math.max(0, Math.round(Number(c.left) || 0)));
    });
    const now = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('updatedAt', now);
    return { updatedAt: now };
  } finally {
    lock.releaseLock();
  }
}
