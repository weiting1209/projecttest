/**
 * 金黃鍋貼 · 訂單接收（Google Apps Script Web App）
 *
 * 功能
 *  - 接收網站結帳表單送來的訂單（POST / JSON）
 *  - 以「伺服器端價目表」重新計算金額，避免有人竄改前端價格
 *  - 寫入「訂單」「訂單明細」兩張工作表（自動建立、自動加標題與狀態下拉選單）
 *  - 訂單編號：GT + 日期 + 當日流水號（LockService 防止同時下單撞號）
 *  - 重複送出防護（requestId）、基本欄位檢查、試算表公式注入防護
 *  - 新訂單 Email 通知（選用）
 *
 * 使用方式請看同資料夾的 README.md
 */

const CONFIG = {
  SPREADSHEET_ID: '',   // 留空 = 使用「此腳本綁定的試算表」（從試算表 → 擴充功能 → Apps Script 開啟）
  NOTIFY_EMAIL: '',     // 新訂單通知信收件人；留空 = 不寄信（多人用逗號分隔）
  TIMEZONE: 'Asia/Taipei',
  ORDER_SHEET: '訂單',
  ITEM_SHEET: '訂單明細',
  STATUSES: ['待處理', '製作中', '已完成', '已取消'],
  MAX_ITEMS: 30,        // 一張訂單最多幾種品項
  MAX_QTY: 50           // 單一品項最多幾份
};

/** 價目表（與網站 index.html 的 MENU_DATA / 禮盒一致；改價請兩邊一起改） */
const MENU = {
  gt1: { name: '原味豬肉鍋貼', price: 65 },
  gt2: { name: '韭菜豬肉鍋貼', price: 65 },
  gt3: { name: '泡菜豬肉鍋貼', price: 70 },
  gt4: { name: '蝦仁鮮肉鍋貼', price: 80 },
  sj1: { name: '高麗菜水餃', price: 55 },
  sj2: { name: '韭菜豬肉水餃', price: 55 },
  sj3: { name: '鮮蝦水餃', price: 70 },
  sj4: { name: '素食水餃', price: 50 },
  sp1: { name: '招牌貢丸湯', price: 50 },
  sp2: { name: '蛤蜊豆腐湯', price: 65 },
  sp3: { name: '珍珠奶茶', price: 65 },
  sp4: { name: '冬瓜檸檬茶', price: 55 },
  sd1: { name: '涼拌小黃瓜', price: 40 },
  sd2: { name: '滷海帶', price: 35 },
  sd3: { name: '皮蛋豆腐', price: 55 },
  sd4: { name: '紅油抄手', price: 60 },
  gift1: { name: '招牌鍋貼禮盒 50 入', price: 680 },
  gift2: { name: '節慶豪華禮盒 100 入', price: 1280 }
};

const DINE = { 'dine-in': '內用', 'takeout': '外帶', 'delivery': '宅配' };

const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '用餐方式', '姓名', '電話', '送貨地址', '備註', '品項摘要', '總金額', '請求ID'];
const ITEM_HEADERS = ['訂單編號', '品項代碼', '品名', '單價', '數量', '小計'];

/* ───────────── Web App 入口 ───────────── */

/** 健康檢查：用瀏覽器直接開網址，看到 ok:true 代表部署成功（不會回傳任何訂單資料） */
function doGet() {
  return json_({ ok: true, service: 'jinhuang-orders', time: nowStr_() });
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) throw new UserError_('沒有收到訂單資料');
    let data;
    try { data = JSON.parse(raw); } catch (_) { throw new UserError_('訂單格式錯誤'); }
    return json_(handleOrder_(data));
  } catch (err) {
    if (err instanceof UserError_) return json_({ ok: false, error: err.message });
    console.error(err && err.stack ? err.stack : err);
    return json_({ ok: false, error: '系統忙碌中，請稍後再試或直接來電訂購' });
  }
}

/* ───────────── 核心流程 ───────────── */

function handleOrder_(data) {
  const order = validate_(data);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new UserError_('目前下單人數較多，請稍後再試一次');

  let result;
  try {
    // 重複送出（網路重試、連點）→ 回傳同一張訂單
    const cache = CacheService.getScriptCache();
    const dupKey = 'req_' + order.requestId;
    const dup = order.requestId && cache.get(dupKey);
    if (dup) return Object.assign(JSON.parse(dup), { duplicate: true });

    const ss = getSpreadsheet_();
    const orderSheet = ensureSheet_(ss, CONFIG.ORDER_SHEET, ORDER_HEADERS);
    const itemSheet = ensureSheet_(ss, CONFIG.ITEM_SHEET, ITEM_HEADERS);

    const orderNo = nextOrderNo_();
    const created = nowStr_();
    const summary = order.items.map(function (i) { return i.name + ' ×' + i.qty; }).join('、');

    const row = orderSheet.getLastRow() + 1;
    orderSheet.getRange(row, 1, 1, ORDER_HEADERS.length).setValues([[
      orderNo, created, CONFIG.STATUSES[0], DINE[order.dine], order.name, order.phone,
      order.dine === 'delivery' ? order.addr : '', order.note, summary, order.total, order.requestId
    ]]);
    orderSheet.getRange(row, 3).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUSES, true).setAllowInvalid(false).build()
    );

    const itemRows = order.items.map(function (i) {
      return [orderNo, i.id, i.name, i.price, i.qty, i.price * i.qty];
    });
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
    SpreadsheetApp.flush();

    result = { ok: true, orderNo: orderNo, total: order.total };
    if (order.requestId) cache.put(dupKey, JSON.stringify(result), 6 * 60 * 60);
  } finally {
    lock.releaseLock();
  }

  notify_(result.orderNo, order);   // 寄信失敗不影響下單
  return result;
}

function validate_(d) {
  if (!d || typeof d !== 'object') throw new UserError_('訂單資料不完整');

  const dine = String(d.dine || '');
  if (!DINE[dine]) throw new UserError_('請選擇用餐方式');

  const name = clean_(d.name, 40);
  if (!name) throw new UserError_('請填寫姓名');

  const phone = clean_(d.phone, 20);
  if (phone.replace(/\D/g, '').length < 8) throw new UserError_('請填寫正確的電話');

  const addr = clean_(d.addr, 120);
  if (dine === 'delivery' && !addr) throw new UserError_('宅配請填寫送貨地址');

  if (!Array.isArray(d.items) || !d.items.length) throw new UserError_('購物車是空的');
  if (d.items.length > CONFIG.MAX_ITEMS) throw new UserError_('品項過多，請分成多張訂單');

  const merged = {};
  d.items.forEach(function (it) {
    const id = String(it && it.id || '');
    const qty = Math.floor(Number(it && it.qty));
    if (!MENU[id]) throw new UserError_('有品項已下架，請重新整理頁面再下單');
    if (!(qty >= 1) || qty > CONFIG.MAX_QTY) throw new UserError_('品項數量不正確（每項最多 ' + CONFIG.MAX_QTY + ' 份）');
    merged[id] = (merged[id] || 0) + qty;
  });

  let total = 0;
  const items = Object.keys(merged).map(function (id) {
    const qty = merged[id];
    if (qty > CONFIG.MAX_QTY) throw new UserError_('品項數量不正確（每項最多 ' + CONFIG.MAX_QTY + ' 份）');
    total += MENU[id].price * qty;
    return { id: id, name: MENU[id].name, price: MENU[id].price, qty: qty };
  });

  // 前端算的總額與伺服器不同 → 通常是價格調整後頁面還沒更新
  if (d.clientTotal != null && Number(d.clientTotal) !== total) {
    throw new UserError_('價格已更新，請重新整理頁面後再下單');
  }

  return {
    dine: dine, name: name, phone: phone, addr: addr, note: clean_(d.note, 200),
    items: items, total: total, requestId: clean_(d.requestId, 64)
  };
}

/* ───────────── 工具 ───────────── */

function nextOrderNo_() {
  const props = PropertiesService.getScriptProperties();
  const day = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');
  const key = 'seq_' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return 'GT' + day + '-' + ('000' + n).slice(-3);
}

function getSpreadsheet_() {
  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表：請從試算表開啟 Apps Script，或填入 CONFIG.SPREADSHEET_ID');
  return ss;
}

/** 工作表不存在就建立，並設定標題列、凍結、純文字欄位（避免電話 0 被吃掉、避免公式注入） */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#3D2B1F').setFontColor('#F7F0E6');
  sh.setFrozenRows(1);
  const textCols = name === CONFIG.ORDER_SHEET ? [1, 2, 3, 4, 5, 6, 7, 8, 9, 11] : [1, 2, 3];
  textCols.forEach(function (c) { sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('@'); });
  sh.autoResizeColumns(1, headers.length);
  return sh;
}

function notify_(orderNo, o) {
  if (!CONFIG.NOTIFY_EMAIL) return;
  try {
    const lines = o.items.map(function (i) { return '・' + i.name + ' ×' + i.qty + '　NT$' + (i.price * i.qty); });
    const body = [
      '訂單編號：' + orderNo,
      '用餐方式：' + DINE[o.dine],
      '姓名：' + o.name,
      '電話：' + o.phone,
      o.dine === 'delivery' ? '地址：' + o.addr : '',
      o.note ? '備註：' + o.note : '',
      '',
      lines.join('\n'),
      '',
      '合計：NT$' + o.total
    ].filter(function (l, i, a) { return l !== '' || (a[i - 1] !== '' && i > 0); }).join('\n');
    MailApp.sendEmail({ to: CONFIG.NOTIFY_EMAIL, subject: '[新訂單] ' + orderNo + '　NT$' + o.total, body: body });
  } catch (err) {
    console.error('通知信寄送失敗：' + err);
  }
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, max);
}
function nowStr_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'); }
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function UserError_(msg) { this.message = msg; this.name = 'UserError'; }
UserError_.prototype = Object.create(Error.prototype);

/* ───────────── 管理用 ───────────── */

/** 第一次部署前，在編輯器選這個函式按「執行」：建立工作表並完成授權 */
function setup() {
  const ss = getSpreadsheet_();
  ensureSheet_(ss, CONFIG.ORDER_SHEET, ORDER_HEADERS);
  ensureSheet_(ss, CONFIG.ITEM_SHEET, ITEM_HEADERS);
  console.log('完成：已建立「' + CONFIG.ORDER_SHEET + '」與「' + CONFIG.ITEM_SHEET + '」工作表');
}

/** 試算表上方多一個選單，方便手動初始化 */
function onOpen() {
  try { SpreadsheetApp.getUi().createMenu('金黃鍋貼').addItem('初始化工作表', 'setup').addToUi(); } catch (_) {}
}
