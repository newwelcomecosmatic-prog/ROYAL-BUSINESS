// ROYAL BUSINESS - one workbook: Orders + PostEx auto-sync + Rider bulk status + Stock
const TZ = 'Asia/Karachi', DATA = 'Data', IT = 'Items', SHOP = 'Shop Items', PR = 'Products', SIN = 'StockIn', DL = 'Deals', MV = 'Moves', PRICE = 'PRODUCT';
const PT = 'Parties', BL = 'Bills', BI = 'BillItems', PY = 'Payments', VD = 'Vendors', VP = 'VendorPay', EX = 'Expenses';
const VB = 'VendorBills', BT = 'Batches', BX = 'BatchIssues';
const DEFAULT_CATS = ['Salary','Rent','Paper','Packaging','Electricity Bill','Gas Bill','Internet Bill','Ads Bill','Website Bill','Transport','Courier Charges','Maintenance','Other'];
const PX = 'https://api.postex.pk/services/integration/api/order/';
const FINAL = /^(delivered|returned|cancel+ed|cancal)$/i, RETURNED = /^(returned|cancel+ed|cancal)$/i;
const RIDER_ST = ['Booked', 'Out for Delivery', 'Pending', 'Delivered', 'Returned', 'Cancelled'], RIDER_OK = ['Out for Delivery', 'Pending', 'Delivered', 'Returned'], SL = 'StatusLog';
const HEADS = {};
HEADS[DATA] = ['ORDER NO','DATE','NAME','NUMBER','CITY','COD','STATUS','TRACKING ID','LAST SYNCED','ADDRESS','ORDER TAKER','ITEMS','AMOUNT','D.C','ADVANCE','TOTAL QTY','DELIVERY'];
HEADS[IT] = ['Order #','Date','Month','Year','Product','Qty','Unit Price (Rs)','Amount (Rs)','Shop Item','Total Units'];
HEADS[SHOP] = ['DATE','ORDER NO','PRODUCT','QTY','PRICE','AMOUNT'];
HEADS[PR] = ['Product','Opening','Purchased','Sold','Returned','Stock','Cost Price'];
HEADS[SIN] = ['Date','Product','Qty','Rate','Supplier','Note','PayType'];
HEADS[PT] = ['Party','Opening','Phone','Added'];
HEADS[BL] = ['Bill','Date','Party','Total','Paid','Status','CreatedBy','PrevBal'];
HEADS[BI] = ['Bill','Product','Qty','Rate','Amount','Shop Item'];
HEADS[PY] = ['Date','Party / Source','Amount','Receipt','By','BalanceAfter','Kind','Note'];
HEADS[VD] = ['Vendor','Opening','Phone','Added'];
HEADS[VP] = ['Date','Vendor','Amount','Voucher','By','BalanceAfter','Note'];
HEADS[EX] = ['Date','Category','Amount','Note','ID','By'];
HEADS[SL] = ['Time','Role','Order No','Status'];
HEADS[VB] = ['Bill No','Date','Vendor','Product','Qty','Rate','Amount','Note'];
HEADS[BT] = ['Batch','Date','Vendor','Product','Planned','Status','BillNo','BillAmount','PackCost','CostPerUnit','Note','BillType'];
HEADS[BX] = ['Batch','Date','Item','Qty','Rate','Amount'];
HEADS[DL] = ['Deal','Product','Qty'];
HEADS[MV] = ['Date','Order No','Product','Qty','Type','Note'];
HEADS[PRICE] = ['Name','Price'];
const PRODUCTS = [["ACNE TUBE", ""], ["ANTI ACNE SERUM", ""], ["BASE", 800], ["BEAUTY CREAM", 800], ["BOTTLE H/A", 800], ["CAPSULE", 100], ["CAPSULE LIPSTICK", 800], ["COMPLETE SKIN CARE DEAL", 2500], ["DEAL 1500", 1500], ["DEAL 2000", 2000], ["DEAL 2000 POLISH", 2000], ["DEAL 2000 SERUM", 2000], ["DEAL 2500", 2500], ["DEAL 3500", 3500], ["DEAL FACE", 1500], ["DEAL H&F", 1500], ["DEAL PIGMENTATION", 2500], ["DEAL SERUM", 2500], ["FACE DWAI", 200], ["FACIAL JAR", 600], ["H&F DWAI", 400], ["HAIR OIL", 600], ["HAIR REMOVAL LOTION", 700], ["HAND AND FOOT CREAM", 800], ["HEEL CREAM", ""], ["HYDROGEN AMONIA", 250], ["JUNSUI WASH", 700], ["MAHANDI", 100], ["MULTANI MITTI SACHET", 100], ["OLD CREAM", 600], ["PARLOUR SKIN POLISH", ""], ["PLUS DWAI", 400], ["REDNESS", ""], ["ROSE WATER", ""], ["ROYAL BODY WASH", 800], ["SERUM B3", 1000], ["SHAMPOO", ""], ["SKIN POLISH", 800], ["SKIN POLISH DEVELPOR", ""], ["SKIN POLISH SACHET", 100], ["SUN BLOCK", 700], ["URGENT FACIAL SACHET", 100], ["WASH L", 700], ["WASH S", 300], ["WAX", ""]];
const P = k => PropertiesService.getScriptProperties().getProperty(k) || '';
const sh_ = n => SpreadsheetApp.getActive().getSheetByName(n);
const fd = d => d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d || '');
const norm_ = s => String(s || '').replace(/\s+/g, ' ').trim().toUpperCase();
const dig_ = s => String(s == null ? '' : s).replace(/\D/g, '');
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
// Message: popup if possible, always also written to the Execution log (works from the phone editor)
function say_(m) { Logger.log(m); try { SpreadsheetApp.getUi().alert(m); } catch (e) {} }
function rows_(n, nc) { const s = sh_(n), l = s ? s.getLastRow() : 0; return l > 1 ? s.getRange(2, 1, l - 1, nc).getValues() : []; }
// One workbook, four apps. Each app has its own PIN (Script Properties) and the backend only allows that app's actions.
const ROLE_PINS = [['APP_PIN', 'manager'], ['CEO_PIN', 'ceo'], ['TRACK_PIN', 'track'], ['RIDER_PIN', 'rider']];
const ALLOW = { track: ['search', 'sync'], rider: ['riderList', 'riderSet', 'scan'], ceo: ['overview', 'search', 'stock', 'parties', 'vendors', 'ledger', 'cash', 'batches', 'batchDetail', 'export'] };
function roleOf_(pin) { for (let i = 0; i < ROLE_PINS.length; i++) if (P(ROLE_PINS[i][0]) && String(pin) === P(ROLE_PINS[i][0])) return ROLE_PINS[i][1]; Utilities.sleep(1500); throw new Error('Wrong PIN'); }
function chk_(pin) { return roleOf_(pin); }
function locked_(fn) { const l = LockService.getScriptLock(); l.waitLock(30000); try { return fn(); } finally { l.releaseLock(); } }

/* ---------- menu & setup ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Royal Business')
    .addItem('1. Create tabs (once)', 'setupTabs').addItem('2. Set PostEx token', 'setToken').addItem('3. Set app PIN', 'setPin')
    .addItem('4. Test PostEx connection', 'testPostEx').addSeparator()
    .addItem('Sync PostEx now', 'syncMenu').addItem('Enable auto-sync (30 min)', 'enableAuto').addItem('Disable auto-sync', 'disableAuto').addToUi();
}
function setupTabs() {   // safe to run again: never deletes data; each tab is done separately so one failure does not stop the rest
  const ss = SpreadsheetApp.getActive(), done = [], failed = [];
  Object.keys(HEADS).forEach(n => {
    try {
      const s = ss.getSheetByName(n) || ss.insertSheet(n), h = HEADS[n];
      s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF'); s.setFrozenRows(1); done.push(n);
    } catch (e) { failed.push(n + ': ' + e.message); }
    SpreadsheetApp.flush();
  });
  try { setupFormulas(); } catch (e) { failed.push('Products formulas: ' + e.message); }
  try { if (sh_(PRICE) && sh_(PRICE).getLastRow() < 2) sh_(PRICE).getRange(2, 1, PRODUCTS.length, 2).setValues(PRODUCTS); } catch (e) { failed.push('PRODUCT prices: ' + e.message); }
  try { const d = ss.getSheetByName('Sheet1'); if (d && d.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(d); } catch (e) {}
  say_('Tabs done: ' + done.length + '/' + Object.keys(HEADS).length + (failed.length ? '\nFAILED:\n' + failed.join('\n') : '\nAll OK. Next: Products, Deals, token and PIN.'));
}
function setupFormulas() {   // Products tab: Purchased / Sold / Returned / Stock formulas
  const p = sh_(PR), f = [];
  for (let r = 2; r <= 300; r++) f.push([
    '=IF(A' + r + '="","",SUMIF(StockIn!B:B,A' + r + ',StockIn!C:C))', '=IF(A' + r + '="","",SUMIFS(Moves!D:D,Moves!C:C,A' + r + ',Moves!E:E,"OUT"))',
    '=IF(A' + r + '="","",SUMIFS(Moves!D:D,Moves!C:C,A' + r + ',Moves!E:E,"IN"))', '=IF(A' + r + '="","",B' + r + '+C' + r + '+E' + r + '-D' + r + ')']);
  p.getRange(2, 3, 299, 4).setFormulas(f);
}
function ask_(title, key) {
  const ui = SpreadsheetApp.getUi(), r = ui.prompt(title, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim()) { PropertiesService.getScriptProperties().setProperty(key, r.getResponseText().trim()); ui.alert('Saved.'); }
}
// EASY WAY (phone / editor): type your values between the quotes below, press Run, then delete the values again.
function saveSecrets() {
  const POSTEX_TOKEN = '', APP_PIN = '';
  const p = PropertiesService.getScriptProperties();
  if (POSTEX_TOKEN) p.setProperty('POSTEX_TOKEN', POSTEX_TOKEN);
  if (APP_PIN) p.setProperty('APP_PIN', APP_PIN);
  say_('Saved. PostEx token: ' + (P('POSTEX_TOKEN') ? 'yes' : 'NO') + ' | PIN: ' + (P('APP_PIN') ? 'yes' : 'NO'));
}
function setToken() { ask_('PostEx merchant API token', 'POSTEX_TOKEN'); }
function setPin() { ask_('App PIN (give it to your team)', 'APP_PIN'); }
function enableAuto() { disableAuto(); ScriptApp.newTrigger('autoTick').timeBased().everyMinutes(30).create(); say_('Auto-sync ON (every 30 min).'); }
function disableAuto() { ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'autoTick') ScriptApp.deleteTrigger(t); }); }
function autoTick() { try { syncPostEx_(); } catch (e) { Logger.log(e.message); } }
function syncMenu() { try { const r = syncPostEx_(); say_('Tracking found: ' + r.tracked + '\nStatus updated: ' + r.updated + (r.diag ? '\n\n' + r.diag : '')); } catch (e) { say_(e.message); } }
function testPostEx() {
  const e = fd(new Date()), s = fd(new Date(Date.now() - 7 * 864e5)), r = pxGet_('v1/get-all-order', 'orderStatusId=0&startDate=' + s + '&endDate=' + e);
  say_('HTTP ' + r.code + '\n' + String(r.body).substring(0, 700));
}

/* ---------- PostEx ---------- */
function pxGet_(path, qs) {
  const tok = P('POSTEX_TOKEN'); if (!tok) throw new Error('Set the PostEx token first (menu).');
  let code = 0, body = '';
  try { const r = UrlFetchApp.fetch(PX + path + '?' + qs, { method: 'get', headers: { token: tok }, muteHttpExceptions: true }); code = r.getResponseCode(); body = r.getContentText(); } catch (e) { body = e.message; }
  let json = null; try { json = JSON.parse(body); } catch (e) {}
  return { ok: code === 200 && json, code: code, body: body, json: json };
}
const pxList_ = j => Array.isArray(j) ? j : (j && (j.dist || j.data)) || [];
function syncPostEx_() {
  return locked_(() => {
    let stage = 'read';
    try {
    const sh = sh_(DATA), n = sh.getLastRow(); if (n < 2) return { tracked: 0, updated: 0 };
    const v = sh.getRange(2, 1, n - 1, 17).getValues(); let tracked = 0, updated = 0, diag = '';
    const need = v.filter(x => x[0] && x[16] === 'POSTEX' && !String(x[7]).trim());
    if (need.length) {   // A) tracking number by order number
      stage = 'tracking';
      const m = {}, end = new Date(), ts = need.map(x => x[1] instanceof Date ? x[1].getTime() : NaN).filter(t => !isNaN(t));
      // start 3 days earlier: PostEx booking date can differ from the bill date
      for (let a = new Date((ts.length ? Math.min.apply(null, ts) : end.getTime() - 30 * 864e5) - 3 * 864e5); a <= end; a = new Date(a.getTime() + 30 * 864e5)) {
        const b = new Date(Math.min(a.getTime() + 29 * 864e5, end.getTime()));
        const r = pxGet_('v1/get-all-order', 'orderStatusId=0&startDate=' + fd(a) + '&endDate=' + fd(b));
        if (!r.ok) { diag = 'get-all-order (' + fd(a) + ' to ' + fd(b) + ') HTTP ' + r.code + ': ' + String(r.body).substring(0, 300); continue; }
        pxList_(r.json).forEach(it => { const ref = dig_(it.orderRefNumber || it.customerOrderId || it.orderRef), tn = it.trackingNumber || ''; if (ref && tn) m[ref] = { tn: tn, st: it.transactionStatus || it.orderStatus || '' }; });
      }
      v.forEach(x => { const f = m[dig_(x[0])]; if (f && x[16] === 'POSTEX' && !String(x[7]).trim()) { x[7] = f.tn; if (f.st) x[6] = f.st; x[8] = new Date(); tracked++; } });
    }
    stage = 'status';
    const live = v.filter(x => String(x[7]).trim() && !FINAL.test(String(x[6]).trim())), idx = {};   // B) status of live parcels
    live.forEach(x => idx[String(x[7]).trim()] = x);
    const ids = Object.keys(idx);
    for (let i = 0; i < ids.length; i += 10) {
      const b = ids.slice(i, i + 10);
      let r = pxGet_('v1/track-bulk-order', 'TrackingNumbers=' + encodeURIComponent(b.join(',')));
      if (!r.ok) r = pxGet_('v1/track-bulk-order', b.map(t => 'TrackingNumbers=' + encodeURIComponent(t)).join('&'));
      if (!r.ok) { diag = 'track-bulk-order HTTP ' + r.code + ': ' + String(r.body).substring(0, 200); continue; }
      pxList_(r.json).forEach(it => {
        const tn = it.trackingNumber || (it.trackingResponse && it.trackingResponse.trackingNumber), st = (it.trackingResponse && it.trackingResponse.transactionStatus) || it.transactionStatus;
        if (tn && idx[tn] && st) { idx[tn][6] = st; idx[tn][8] = new Date(); updated++; }
      });
    }
    stage = 'write';
    sh.getRange(2, 7, v.length, 3).setValues(v.map(x => [x[6], x[7], x[8]]));
    stage = 'returns';
    applyReturns_();
    return { tracked: tracked, updated: updated, diag: tracked + updated === 0 ? diag : '' };
    } catch (e) { throw new Error('[' + stage + '] ' + e.message); }
  });
}
// Returned / cancelled orders: put their stock back (once per order)
function applyReturns_() {
  const dv = rows_(DATA, 17), mv = rows_(MV, 6), outO = {}, inO = {}, ret = {};
  mv.forEach(x => { if (!+x[1]) return; if (x[4] === 'OUT') outO[+x[1]] = 1; if (x[4] === 'IN') inO[+x[1]] = 1; });
  dv.forEach(x => { if (RETURNED.test(String(x[6]).trim()) && outO[+x[0]] && !inO[+x[0]]) ret[+x[0]] = 1; });
  const d = new Date(), add = mv.filter(x => +x[1] && x[4] === 'OUT' && ret[+x[1]]).map(x => [d, x[1], x[2], x[3], 'IN', 'Returned']);
  if (add.length) sh_(MV).getRange(sh_(MV).getLastRow() + 1, 1, add.length, 6).setValues(add);
}

/* ---------- catalog / stock ---------- */
function catalog_() {
  const c = {}, prod = {};
  rows_(PR, 1).forEach(x => { if (x[0]) { prod[norm_(x[0])] = { n: String(x[0]).trim() }; c[norm_(x[0])] = { n: String(x[0]).trim(), deal: 0 }; } });
  const dl = {};
  rows_(DL, 3).forEach(x => { if (x[0] && x[1]) (dl[norm_(x[0])] = dl[norm_(x[0])] || { n: String(x[0]).trim(), deal: 1, comps: [] }).comps.push({ p: (prod[norm_(x[1])] || { n: String(x[1]).trim() }).n, q: +x[2] || 0 }); });
  Object.keys(dl).forEach(k => c[k] = dl[k]);
  return c;
}
function stockMap_() { SpreadsheetApp.flush(); const m = {}; rows_(PR, 6).forEach(x => { if (x[0]) m[String(x[0]).trim()] = +x[5] || 0; }); return m; }

/* ---------- app API ---------- */
function doGet() { return out_({ ok: 1 }); }
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents), role = roleOf_(b.pin); if (role !== 'manager' && (ALLOW[role] || []).indexOf(b.a) < 0) throw new Error('Not allowed in this app'); const A = {
      state: () => state_(), save: () => saveOrder_(b.o), search: () => search_(b.q), riderList: () => riderList_(), riderSet: () => riderSet_(b.orders, b.status, role), scan: () => scan_(b.code), overview: () => overview_(),
      sync: () => syncPostEx_(), stock: () => stock_(), stockIn: () => stockIn_(b), addProduct: () => addProduct_(b), export: () => exportXlsx_(b.from, b.to),
      parties: () => ({ rows: partyBals_() }), addParty: () => addNamed_(PT, b), partyBill: () => partyBill_(b), cancelBill: () => cancelBill_(b.id), receive: () => receive_(b), ledger: () => ledger_(b),
      vendors: () => ({ rows: vendorBals_() }), addVendor: () => addNamed_(VD, b), vendorBill: () => vendorBill_(b), batches: () => batches_(), batchNew: () => batchNew_(b), batchIssue: () => batchIssue_(b), batchReceive: () => batchReceive_(b), batchBill: () => batchBill_(b), batchDetail: () => batchDetail_(b.id), delVendorBill: () => delVendorBill_(b), vendorPay: () => vendorPay_(b), expense: () => expense_(b), delExpense: () => delExpense_(b.id), cash: () => cash_(), setCash: () => { const p = PropertiesService.getScriptProperties(); p.setProperty('OPENING_CASH', String(+b.amount || 0)); p.setProperty('CASH_START', /^\d{4}-\d{2}-\d{2}$/.test(b.from || '') ? b.from : ''); return { ok: 1 }; } };
    if (!A[b.a]) throw new Error('Unknown action');
    return out_(A[b.a]());
  } catch (err) { return out_({ error: err.message }); }
}
const mapRow_ = x => ({ no: x[0], date: fd(x[1]), name: x[2], phone: String(x[3]), city: x[4], cod: x[5], status: x[6], tracking: x[7], synced: x[8] instanceof Date ? Utilities.formatDate(x[8], TZ, 'dd MMM HH:mm') : '', delivery: x[16],
  address: x[9], taker: x[10], items: x[11], amount: x[12], dc: x[13], adv: x[14], qty: x[15] });
function state_() {
  const v = rows_(DATA, 17).filter(x => x[0]), t = fd(new Date()); let mx = 0, taker = '', dc = 350, today = 0, cod = 0, open = 0;
  v.forEach(x => { if (+x[0] > mx) mx = +x[0]; taker = x[10] || taker; if (x[13] !== '') dc = x[13]; if (fd(x[1]) === t) { today++; cod += +x[5] || 0; } if (!FINAL.test(String(x[6]).trim())) open++; });
  return { next: mx ? mx + 1 : '', taker: taker, dc: dc, today: t, kpi: { today: today, cod: cod, open: open },
    products: rows_(PRICE, 2).filter(r => String(r[0]).trim()).map(r => ({ n: String(r[0]).trim(), p: r[1] })) };
}
function saveOrder_(o) {
  return locked_(() => {
    const no = Number(o.orderNo), items = (o.items || []).map(i => ({ p: norm_(i.p), q: +i.q, r: +i.r, ns: !!i.ns })).filter(i => i.p && i.q > 0);
    if (!no) throw new Error('Enter the order number'); if (!items.length) throw new Error('Add at least one item');
    const ds = sh_(DATA), n = ds.getLastRow();
    if (n > 1 && ds.getRange(2, 1, n - 1, 1).getValues().some(x => +x[0] === no)) throw new Error('Order ' + no + ' already exists');
    const pl = plan_(items), mv = pl.mv, warn = pl.warn;
    const d = Utilities.parseDate(o.date, TZ, 'yyyy-MM-dd'); let num = String(o.number || '').trim(); if (/^\d{10}$/.test(num)) num = '0' + num;
    const amt = items.reduce((s, i) => s + i.q * i.r, 0), dc = +o.dc || 0, adv = +o.adv || 0, pay = amt + dc - adv, del = o.delivery === 'RIDER' ? 'RIDER' : 'POSTEX';
    const r = n + 1; ds.getRange(r, 4).setNumberFormat('@');
    ds.getRange(r, 1, 1, 17).setValues([[no, d, o.name, num, o.city, pay, del === 'RIDER' ? 'Booked' : '', '', '', o.address, o.taker, items.map(i => i.p + ' x ' + i.q + ', ').join(''), amt, dc, adv, items.reduce((s, i) => s + i.u, 0), del]]);
    const mo = Utilities.formatDate(d, TZ, 'MMMM'), yr = +Utilities.formatDate(d, TZ, 'yyyy'), si = sh_(IT);
    si.getRange(si.getLastRow() + 1, 1, items.length, 10).setValues(items.map(i => [no, d, mo, yr, i.p, i.q, i.r, i.q * i.r, i.ns ? 'Y' : '', i.u]));
    const ns = items.filter(i => i.ns), sp = sh_(SHOP);
    if (ns.length) sp.getRange(sp.getLastRow() + 1, 1, ns.length, 6).setValues(ns.map(i => [d, no, i.p, i.q, i.r, i.q * i.r]));
    if (mv.length) sh_(MV).getRange(sh_(MV).getLastRow() + 1, 1, mv.length, 6).setValues(mv.map(m => [d, no, m.p, m.q, 'OUT', 'Bill']));
    return { orderNo: no, payable: pay, warn: warn, next: no + 1 };
  });
}
function search_(q) {
  q = String(q || '').trim(); const dq = dig_(q), lq = q.toLowerCase(); let r = rows_(DATA, 17).filter(x => x[0]);
  if (q) r = r.filter(x => (dq && (String(x[0]) === dq || dig_(x[7]) === dq || (dq.length >= 7 && dig_(x[3]).indexOf(dq) >= 0))) || String(x[2]).toLowerCase().indexOf(lq) >= 0);
  return { rows: r.sort((a, b) => b[0] - a[0]).slice(0, 100).map(mapRow_) };
}
function riderList_() {
  const all = rows_(DATA, 17).filter(x => x[0] && x[16] === 'RIDER'), t = fd(new Date()); let n = 0, cod = 0;
  all.forEach(x => { if (/^delivered$/i.test(String(x[6]).trim()) && x[8] instanceof Date && fd(x[8]) === t) { n++; cod += +x[5] || 0; } });
  return { statuses: RIDER_ST, today: { delivered: n, cod: cod }, rows: all.filter(x => !FINAL.test(String(x[6]).trim())).sort((a, b) => b[0] - a[0]).map(mapRow_) };
}
function scan_(code) {   // rider scans a parcel label (order number) -> parcel details
  const c = dig_(code); if (!c) throw new Error('Could not read the code');
  const x = rows_(DATA, 17).find(r => r[0] && (dig_(r[0]) === c || (dig_(r[7]) && dig_(r[7]) === c)));
  if (!x) throw new Error('Order ' + c + ' not found'); if (x[16] !== 'RIDER') throw new Error('Order ' + x[0] + ' is a PostEx parcel, not a rider parcel');
  return { order: mapRow_(x) };
}
function riderSet_(orders, status, role) {
  const st = RIDER_ST.filter(s => s.toLowerCase() === String(status).toLowerCase())[0]; if (!st) throw new Error('Invalid status');
  if (role === 'rider' && RIDER_OK.indexOf(st) < 0) throw new Error('Not allowed');
  const want = {}; (orders || []).forEach(o => { const k = dig_(o); if (k) want[k] = 1; });
  return locked_(() => {
    const sh = sh_(DATA), v = rows_(DATA, 17), done = [], skipped = [], now = new Date();
    v.forEach((x, i) => { const k = dig_(x[0]); if (!want[k]) return; delete want[k];
      if (x[16] !== 'RIDER') { skipped.push(k + ' (PostEx parcel)'); return; }
      sh.getRange(i + 2, 7).setValue(st); sh.getRange(i + 2, 9).setValue(now); done.push(k); });
    Object.keys(want).forEach(k => skipped.push(k + ' (not found)'));
    const lg = sh_(SL); if (lg && done.length) lg.getRange(lg.getLastRow() + 1, 1, done.length, 4).setValues(done.map(k => [now, role || 'manager', Number(k), st]));
    applyReturns_();
    return { done: done.length, skipped: skipped };
  });
}
function overview_() {   // CEO dashboard numbers
  const v = rows_(DATA, 17).filter(x => x[0]), t = fd(new Date()), mo = t.slice(0, 7), st = {}; let tn = 0, tc = 0, mn = 0, mc = 0;
  v.forEach(x => { const d = fd(x[1]), c = +x[5] || 0; if (d === t) { tn++; tc += c; } if (d.slice(0, 7) === mo) { mn++; mc += c; const s = String(x[6]).trim() || 'Not booked'; st[s] = (st[s] || 0) + 1; } });
  const cs = cash_();
  return { today: t, tn: tn, tc: tc, mn: mn, mc: mc, status: st, inHand: cs.inHand, cashIn: cs.month.i, cashOut: cs.month.o,
    receivable: partyBals_().reduce((s, x) => s + (x.bal > 0 ? x.bal : 0), 0), payable: vendorBals_().reduce((s, x) => s + (x.bal > 0 ? x.bal : 0), 0),
    low: stock_().rows.filter(x => x.stock <= 5).sort((a, b) => a.stock - b.stock).slice(0, 20) };
}
function stock_() { const m = stockMap_(); return { rows: rows_(PR, 7).filter(x => x[0]).map(x => ({ n: x[0], stock: +x[5] || 0, cost: x[6] })), total: Object.keys(m).length }; }
function stockIn_(b) {   // stock only: adds quantity, no vendor bill. Tick "cash" for a market purchase paid in cash.
  const q = +b.qty; if (!(q > 0)) throw new Error('Enter a quantity');
  const c = catalog_()[norm_(b.product)]; if (!c || c.deal) throw new Error('Product not found in Products');
  const cash = !!b.cash; if (cash && !(+b.rate > 0)) throw new Error('Enter the rate for a cash purchase');
  return locked_(() => {
    const s = sh_(SIN); s.getRange(s.getLastRow() + 1, 1, 1, 7).setValues([[pd_(b.date), c.n, q, +b.rate || '', String(b.vendor || '').trim(), b.note || '', cash ? 'CASH' : 'STOCK']]);
    if (+b.rate) { const p = sh_(PR), r = rows_(PR, 1); r.forEach((x, i) => { if (norm_(x[0]) === norm_(c.n)) p.getRange(i + 2, 7).setValue(+b.rate); }); }
    return { ok: 1 };
  });
}
function addProduct_(b) {
  const nm = String(b.name || '').trim(); if (!nm) throw new Error('Enter a product name');
  if (catalog_()[norm_(nm)]) throw new Error('Already exists');
  return locked_(() => {
    const p = sh_(PR), v = p.getRange(2, 1, 299, 1).getValues(), i = v.findIndex(x => !String(x[0]).trim()); if (i < 0) throw new Error('Product limit (299) reached');
    p.getRange(i + 2, 1, 1, 2).setValues([[nm, +b.opening || 0]]); if (+b.cost) p.getRange(i + 2, 7).setValue(+b.cost); return { ok: 1 };
  });
}

/* ================= PARTIES / VENDORS / CASH ================= */
const eq_ = (a, b) => norm_(a) === norm_(b);
const pd_ = s => Utilities.parseDate(s || fd(new Date()), TZ, 'yyyy-MM-dd');
const names_ = n => rows_(n, 1).map(x => String(x[0]).trim()).filter(Boolean);
function nextId_(n, pre, col) { let mx = 1000; rows_(n, col + 1).forEach(x => { const m = String(x[col]).match(/(\d+)$/); if (m && +m[1] > mx) mx = +m[1]; }); return pre + (mx + 1); }
function plan_(items) {   // stock movements + low-stock warnings for a list of items (deals expand to their products)
  const cat = catalog_(), st = stockMap_(), bad = [], mv = [], need = {}, warn = [];
  items.forEach(i => {
    if (i.ns) { i.u = i.q; return; }
    const c = cat[i.p]; if (!c) { bad.push(i.p); return; }
    const parts = c.deal ? c.comps.map(x => ({ p: x.p, q: x.q * i.q })) : [{ p: c.n, q: i.q }];
    i.u = c.deal ? (c.comps.reduce((t, x) => t + x.q, 0) || 1) * i.q : i.q;
    parts.forEach(m => { mv.push(m); need[m.p] = (need[m.p] || 0) + m.q; });
  });
  if (bad.length) throw new Error('Not in Products/Deals: ' + bad.join(', ') + '. Tick "Shop item" or add it under Stock.');
  Object.keys(need).forEach(p => { if ((st[p] || 0) < need[p]) warn.push(p + ': stock ' + (st[p] || 0) + ', needed ' + need[p]); });
  return { mv: mv, warn: warn };
}
function sortEv_(ev, f) { ev.sort((a, b) => a.d < b.d ? -1 : a.d > b.d ? 1 : a.k < b.k ? -1 : 1); let bal = 0; ev.forEach(e => { bal += f(e); e.bal = bal; }); return ev; }
function pLedger_(name) {   // party owes us: bills (debit) - paid / payments (credit)
  const ev = [];
  rows_(PT, 4).forEach(x => { if (eq_(x[0], name) && +x[1]) ev.push({ d: '', t: 'Opening balance', dr: +x[1], cr: 0, k: '0' }); });
  rows_(BL, 8).forEach(x => { if (eq_(x[2], name) && x[5] !== 'CANCELLED') ev.push({ d: fd(x[1]), t: 'Bill ' + x[0], dr: +x[3] || 0, cr: +x[4] || 0, id: x[0], k: 'A' + x[0] }); });
  rows_(PY, 8).forEach(x => { if (x[6] === 'PARTY' && eq_(x[1], name)) ev.push({ d: fd(x[0]), t: 'Payment received ' + (x[3] || '') + (x[7] ? ' - ' + x[7] : ''), dr: 0, cr: +x[2] || 0, k: 'Z' + x[3] }); });
  return sortEv_(ev, e => e.dr - e.cr);
}
function vLedger_(name) {   // we owe vendor: BILLS (credit) - payments (debit). Received stock does not change this.
  const ev = [], bills = {};
  rows_(VD, 4).forEach(x => { if (eq_(x[0], name) && +x[1]) ev.push({ d: '', t: 'Opening payable', dr: 0, cr: +x[1], k: '0' }); });
  rows_(VB, 8).forEach(x => { if (eq_(x[2], name)) { const id = String(x[0]), o = bills[id] = bills[id] || { d: fd(x[1]), amt: 0, nt: '' }, nt = String(x[7] || ''); o.amt += +x[6] || 0; if (!o.nt) o.nt = /^BATCH:/.test(nt) ? 'batch ' + nt.slice(6).trim() : nt; } });
  Object.keys(bills).forEach(id => ev.push({ d: bills[id].d, t: 'Bill ' + id + (bills[id].nt ? ' - ' + bills[id].nt : ''), dr: 0, cr: bills[id].amt, id: id, k: 'A' + id }));
  rows_(VP, 7).forEach(x => { if (eq_(x[1], name)) ev.push({ d: fd(x[0]), t: 'Payment made ' + (x[3] || '') + (x[6] ? ' - ' + x[6] : ''), dr: +x[2] || 0, cr: 0, k: 'Z' + x[3] }); });
  return sortEv_(ev, e => e.cr - e.dr);
}
function vRecon_(name) {   // billed qty vs received qty per product: d>0 stock pending, d<0 bill pending
  const k = norm_(name), m = {}, g = p => m[norm_(p)] = m[norm_(p)] || { p: String(p).trim(), b: 0, r: 0, rate: 0 };
  rows_(VB, 8).forEach(x => { if (norm_(x[2]) === k && norm_(x[3]) && String(x[7]).indexOf('BATCH:') !== 0) { const o = g(x[3]); o.b += +x[4] || 0; o.rate = +x[5] || o.rate; } });
  rows_(SIN, 7).forEach(x => { if (x[6] === 'VENDOR' && norm_(x[4]) === k) { const o = g(x[1]); o.r += +x[2] || 0; if (!o.rate) o.rate = +x[3] || 0; } });
  const cost = {}; rows_(PR, 7).forEach(x => cost[norm_(x[0])] = +x[6] || 0);
  return Object.keys(m).map(q => { const o = m[q]; o.d = o.b - o.r; if (!o.rate) o.rate = cost[q] || 0; o.val = Math.abs(o.d) * o.rate; return o; }).filter(o => o.d !== 0);
}
const lastBal_ = l => l.length ? l[l.length - 1].bal : 0;
function partyBals_() {
  const b = {}, nm = {};
  rows_(PT, 4).forEach(x => { if (x[0]) { nm[norm_(x[0])] = String(x[0]).trim(); b[norm_(x[0])] = +x[1] || 0; } });
  rows_(BL, 8).forEach(x => { const k = norm_(x[2]); if (k in b && x[5] !== 'CANCELLED') b[k] += (+x[3] || 0) - (+x[4] || 0); });
  rows_(PY, 8).forEach(x => { const k = norm_(x[1]); if (x[6] === 'PARTY' && k in b) b[k] -= +x[2] || 0; });
  return Object.keys(b).map(k => ({ n: nm[k], bal: b[k] })).sort((x, y) => x.n < y.n ? -1 : 1);
}
function vendorBals_() {
  const b = {}, nm = {}, pm = {}, sp = {}, bp = {};
  rows_(VD, 4).forEach(x => { if (x[0]) { nm[norm_(x[0])] = String(x[0]).trim(); b[norm_(x[0])] = +x[1] || 0; } });
  rows_(VB, 8).forEach(x => { const k = norm_(x[2]); if (k in b) { b[k] += +x[6] || 0; if (norm_(x[3]) && String(x[7]).indexOf('BATCH:') !== 0) { const q = k + '|' + norm_(x[3]); pm[q] = (pm[q] || 0) + (+x[4] || 0); } } });
  rows_(SIN, 7).forEach(x => { const k = norm_(x[4]); if (x[6] === 'VENDOR' && k in b) { const q = k + '|' + norm_(x[1]); pm[q] = (pm[q] || 0) - (+x[2] || 0); } });
  rows_(VP, 7).forEach(x => { const k = norm_(x[1]); if (k in b) b[k] -= +x[2] || 0; });
  Object.keys(pm).forEach(q => { const k = q.split('|')[0]; if (pm[q] > 0) sp[k] = (sp[k] || 0) + pm[q]; else if (pm[q] < 0) bp[k] = (bp[k] || 0) - pm[q]; });
  return Object.keys(b).map(k => ({ n: nm[k], bal: b[k], sp: sp[k] || 0, bp: bp[k] || 0 })).sort((x, y) => x.n < y.n ? -1 : 1);
}
function addNamed_(tab, b) {   // new party or vendor
  const nm = String(b.name || '').trim(); if (!nm) throw new Error('Enter a name');
  return locked_(() => {
    if (names_(tab).some(x => eq_(x, nm))) throw new Error('Already exists');
    const s = sh_(tab); s.getRange(s.getLastRow() + 1, 1, 1, 4).setValues([[nm, +b.opening || 0, b.phone || '', new Date()]]); return { ok: 1 };
  });
}
function ledger_(b) {
  const party = b.t === 'party', l = party ? pLedger_(b.name) : vLedger_(b.name);
  return { bal: lastBal_(l), rows: l.slice(-200), recon: party ? [] : vRecon_(b.name) };
}
function partyBill_(b) {
  const party = names_(PT).filter(x => eq_(x, b.party))[0]; if (!party) throw new Error('Select an existing party (add it first).');
  const items = (b.items || []).map(i => ({ p: norm_(i.p), q: +i.q, r: +i.r, ns: !!i.ns })).filter(i => i.p && i.q > 0); if (!items.length) throw new Error('Add at least one item');
  const total = items.reduce((s, i) => s + i.q * i.r, 0), paid = +b.paid || 0; if (paid < 0 || paid > total) throw new Error('Paid amount must be between 0 and the bill total');
  return locked_(() => {
    const pl = plan_(items), prev = lastBal_(pLedger_(party)), id = nextId_(BL, 'B', 0), d = pd_(b.date);
    sh_(BL).getRange(sh_(BL).getLastRow() + 1, 1, 1, 8).setValues([[id, d, party, total, paid, 'ACTIVE', 'App', prev]]);
    sh_(BI).getRange(sh_(BI).getLastRow() + 1, 1, items.length, 6).setValues(items.map(i => [id, i.p, i.q, i.r, i.q * i.r, i.ns ? 'Y' : '']));
    if (pl.mv.length) sh_(MV).getRange(sh_(MV).getLastRow() + 1, 1, pl.mv.length, 6).setValues(pl.mv.map(m => [d, id, m.p, m.q, 'OUT', 'Party bill']));
    return { id: id, prev: prev, total: total, paid: paid, bal: prev + total - paid, warn: pl.warn };
  });
}
function cancelBill_(id) {
  return locked_(() => {
    const s = sh_(BL), v = rows_(BL, 8), i = v.findIndex(x => x[0] === id && x[5] !== 'CANCELLED'); if (i < 0) throw new Error('Bill not found or already cancelled');
    s.getRange(i + 2, 6).setValue('CANCELLED');
    const d = new Date(), add = rows_(MV, 6).filter(x => x[1] === id && x[4] === 'OUT').map(x => [d, id, x[2], x[3], 'IN', 'Bill cancelled']);
    if (add.length) sh_(MV).getRange(sh_(MV).getLastRow() + 1, 1, add.length, 6).setValues(add);
    return { ok: 1 };
  });
}
function receive_(b) {   // PARTY = reduces party balance | ONLINE = weekly courier/PostEx payment (cash in, no balance)
  const amt = +b.amount; if (!(amt > 0)) throw new Error('Enter the amount');
  return locked_(() => {
    const kind = b.kind === 'ONLINE' ? 'ONLINE' : 'PARTY'; let name, after = '';
    if (kind === 'PARTY') { name = names_(PT).filter(x => eq_(x, b.name))[0]; if (!name) throw new Error('Select an existing party'); after = lastBal_(pLedger_(name)) - amt; }
    else name = String(b.name || '').trim() || 'PostEx';
    const id = nextId_(PY, 'R', 3);
    sh_(PY).getRange(sh_(PY).getLastRow() + 1, 1, 1, 8).setValues([[pd_(b.date), name, amt, id, 'App', after, kind, b.note || '']]);
    return { id: id, after: after };
  });
}
function vendorBill_(b) {   // vendor bill = credit only (what we owe). Stock is added separately on the Stock page.
  const v = names_(VD).filter(x => eq_(x, b.vendor))[0]; if (!v) throw new Error('Select an existing vendor');
  const amt = +b.amount; if (!(amt > 0)) throw new Error('Enter the bill amount');
  return locked_(() => {
    const no = String(b.billNo || '').trim() || nextId_(VB, 'VB', 0);
    if (rows_(VB, 8).some(x => eq_(x[2], v) && String(x[0]) === no)) throw new Error('This bill number already exists for this vendor');
    const s = sh_(VB), r = s.getLastRow() + 1; s.getRange(r, 1, 1, 1).setNumberFormat('@');
    s.getRange(r, 1, 1, 8).setValues([[no, pd_(b.date), v, '', 0, 0, amt, b.note || '']]);
    return { total: amt, bal: lastBal_(vLedger_(v)) };
  });
}
function delVendorBill_(b) {
  return locked_(() => {
    const s = sh_(VB), v = rows_(VB, 8); let n = 0;
    for (let i = v.length - 1; i >= 0; i--) if (eq_(v[i][2], b.vendor) && String(v[i][0]) === String(b.billNo)) { s.deleteRow(i + 2); n++; }
    return { deleted: n };
  });
}
function vendorPay_(b) {
  const amt = +b.amount; if (!(amt > 0)) throw new Error('Enter the amount');
  const v = names_(VD).filter(x => eq_(x, b.vendor))[0]; if (!v) throw new Error('Select an existing vendor');
  return locked_(() => {
    const id = nextId_(VP, 'V', 3), after = lastBal_(vLedger_(v)) - amt;
    sh_(VP).getRange(sh_(VP).getLastRow() + 1, 1, 1, 7).setValues([[pd_(b.date), v, amt, id, 'App', after, b.note || '']]);
    return { id: id, after: after };
  });
}
function expense_(b) {
  const cat = String(b.cat || '').trim(), amt = +b.amount; if (!cat || !(amt > 0)) throw new Error('Enter category and amount');
  return locked_(() => { const s = sh_(EX); s.getRange(s.getLastRow() + 1, 1, 1, 6).setValues([[pd_(b.date), cat, amt, b.note || '', nextId_(EX, 'E', 4), 'App']]); return { ok: 1 }; });
}
function delExpense_(id) { return locked_(() => { const i = rows_(EX, 6).findIndex(x => x[4] === id); if (i >= 0) sh_(EX).deleteRow(i + 2); return { ok: 1 }; }); }
function cashRows_() {   // every cash movement. Online bills add NO cash; their weekly payments come in as ONLINE receipts
  const r = [];
  rows_(BL, 8).forEach(x => { if (x[5] !== 'CANCELLED' && +x[4]) r.push({ d: fd(x[1]), io: 'IN', cat: 'Party bill cash', amt: +x[4] }); });
  rows_(PY, 8).forEach(x => r.push({ d: fd(x[0]), io: 'IN', cat: x[6] === 'ONLINE' ? 'Online weekly payments' : 'Party payments', amt: +x[2] || 0 }));
  rows_(VP, 7).forEach(x => r.push({ d: fd(x[0]), io: 'OUT', cat: 'Vendor payments', amt: +x[2] || 0 }));
  rows_(SIN, 7).forEach(x => { if (x[6] === 'CASH') r.push({ d: fd(x[0]), io: 'OUT', cat: 'Market purchases (cash)', amt: (+x[2] || 0) * (+x[3] || 0) }); });
  rows_(EX, 6).forEach(x => r.push({ d: fd(x[0]), io: 'OUT', cat: x[1] || 'Other', amt: +x[2] || 0 }));
  const st = P('CASH_START');   // entries dated before the cash start date stay in the ledgers but do not touch cash
  return st ? r.filter(x => x.d >= st) : r;
}
function cash_() {
  const CR = cashRows_(), t = fd(new Date()), mo = t.slice(0, 7), oc = +P('OPENING_CASH') || 0, ex = rows_(EX, 6);
  const sm = f => CR.filter(f).reduce((s, x) => s + x.amt, 0), cg = io => { const o = {}; CR.forEach(x => { if (x.io === io && x.d.slice(0, 7) === mo) o[x.cat] = (o[x.cat] || 0) + x.amt; }); return o; };
  const cs = {}; DEFAULT_CATS.forEach(c => cs[c] = 1); ex.forEach(r => { if (r[1]) cs[r[1]] = 1; });
  return { open: oc, start: P('CASH_START'), inHand: oc + sm(x => x.io === 'IN') - sm(x => x.io === 'OUT'), today: { i: sm(x => x.d === t && x.io === 'IN'), o: sm(x => x.d === t && x.io === 'OUT') },
    month: { i: sm(x => x.d.slice(0, 7) === mo && x.io === 'IN'), o: sm(x => x.d.slice(0, 7) === mo && x.io === 'OUT'), inCats: cg('IN'), outCats: cg('OUT') }, cats: Object.keys(cs),
    exp: ex.slice(-30).reverse().map(r => ({ id: r[4], date: fd(r[0]), cat: r[1], amt: r[2], note: r[3] })),
    pays: rows_(PY, 8).slice(-15).reverse().map(x => ({ date: fd(x[0]), name: x[1], amt: x[2], kind: x[6], note: x[7] })) };
}

/* ================= BATCHES (vendor makes our product; we send packaging; bill comes in kg at the end) ================= */
function batches_() {
  const rec = {}, pack = {};
  rows_(SIN, 7).forEach(r => { if (r[6] === 'BATCH') rec[String(r[5])] = (rec[String(r[5])] || 0) + (+r[2] || 0); });
  rows_(BX, 6).forEach(r => { pack[String(r[0])] = (pack[String(r[0])] || 0) + (+r[5] || 0); });
  return { rows: rows_(BT, 12).filter(x => x[0]).reverse().map(x => ({ id: String(x[0]), date: fd(x[1]), vendor: x[2], product: x[3], planned: +x[4] || 0, status: x[5], billNo: x[6], bill: +x[7] || 0,
    pack: pack[String(x[0])] || 0, unit: +x[9] || 0, kind: x[11], received: rec[String(x[0])] || 0, note: x[10] })) };
}
function batchGet_(id, open) {
  const v = rows_(BT, 12), i = v.findIndex(x => String(x[0]) === String(id)); if (i < 0) throw new Error('Select a batch');
  if (open && v[i][5] !== 'OPEN') throw new Error('This batch is already closed'); return { row: i + 2, x: v[i] };
}
function batchNew_(b) {
  const v = names_(VD).filter(x => eq_(x, b.vendor))[0]; if (!v) throw new Error('Select an existing vendor');
  const c = catalog_()[norm_(b.product)]; if (!c || c.deal) throw new Error('Product not found in Products'); const q = +b.planned; if (!(q > 0)) throw new Error('Enter the planned quantity');
  return locked_(() => { const id = nextId_(BT, 'BT', 0), s = sh_(BT); s.getRange(s.getLastRow() + 1, 1, 1, 12).setValues([[id, pd_(b.date), v, c.n, q, 'OPEN', '', '', '', '', b.note || '', '']]); return { id: id }; });
}
function batchIssue_(b) {   // packaging (jar, label, box) sent to the vendor: stock goes out, cost is recorded against the batch
  const cat = catalog_(), m = {}, bad = [];
  (b.items || []).forEach(i => { const q = +i.q; if (!norm_(i.p) || !(q > 0)) return; const c = cat[norm_(i.p)]; if (!c || c.deal) bad.push(i.p); else m[c.n] = (m[c.n] || 0) + q; });
  if (bad.length) throw new Error('Not in Products: ' + bad.join(', ')); if (!Object.keys(m).length) throw new Error('Add at least one item');
  return locked_(() => {
    const g = batchGet_(b.id, true), id = String(g.x[0]), d = pd_(b.date), st = stockMap_(), cost = {}, warn = [], noCost = [];
    rows_(PR, 7).forEach(x => cost[String(x[0]).trim()] = +x[6] || 0);
    const bx = Object.keys(m).map(p => { if ((st[p] || 0) < m[p]) warn.push(p + ': stock ' + (st[p] || 0) + ', needed ' + m[p]); if (!cost[p]) noCost.push(p); return [id, d, p, m[p], cost[p] || 0, m[p] * (cost[p] || 0)]; });
    sh_(BX).getRange(sh_(BX).getLastRow() + 1, 1, bx.length, 6).setValues(bx);
    sh_(MV).getRange(sh_(MV).getLastRow() + 1, 1, bx.length, 6).setValues(bx.map(r => [d, id, r[2], r[3], 'OUT', 'Batch packaging']));
    return { packCost: bx.reduce((t, r) => t + r[5], 0), warn: warn, noCost: noCost };
  });
}
function batchReceive_(b) {   // finished goods came back from the vendor (partial deliveries are fine)
  const q = +b.qty; if (!(q > 0)) throw new Error('Enter the quantity received');
  return locked_(() => {
    const g = batchGet_(b.id, true), s = sh_(SIN);
    s.getRange(s.getLastRow() + 1, 1, 1, 7).setValues([[pd_(b.date), g.x[3], q, '', g.x[2], String(g.x[0]), 'BATCH']]); return { ok: 1 };
  });
}
function batchBill_(b) {   // vendor bill in kg arrives at the end: goes to payable, batch closes, cost per piece is worked out
  const kg = +(b.qty || b.kg), rate = +b.rate, no = String(b.billNo || '').trim(), kind = b.kind === 'LABOUR' ? 'LABOUR' : 'MATERIAL'; if (!no || !(kg > 0) || !(rate > 0)) throw new Error('Enter bill no, quantity and rate');
  return locked_(() => {
    const g = batchGet_(b.id, true), id = String(g.x[0]), vendor = g.x[2], product = g.x[3];
    if (rows_(VB, 8).some(x => eq_(x[2], vendor) && String(x[0]) === no)) throw new Error('This bill number already exists for this vendor');
    const rec = rows_(SIN, 7).filter(r => r[6] === 'BATCH' && String(r[5]) === id).reduce((t, r) => t + (+r[2] || 0), 0); if (!(rec > 0)) throw new Error('No finished goods received for this batch yet');
    const pack = rows_(BX, 6).filter(r => String(r[0]) === id).reduce((t, r) => t + (+r[5] || 0), 0), amt = kg * rate, unit = Math.round((amt + pack) / rec * 100) / 100;
    const s = sh_(VB), r = s.getLastRow() + 1; s.getRange(r, 1, 1, 1).setNumberFormat('@');
    s.getRange(r, 1, 1, 8).setValues([[no, pd_(b.date), vendor, product, kg, rate, amt, 'BATCH:' + id + (kind === 'LABOUR' ? ' LABOUR' : '')]]);
    sh_(BT).getRange(g.row, 6, 1, 5).setValues([['CLOSED', no, amt, pack, unit]]); sh_(BT).getRange(g.row, 12).setValue(kind);
    rows_(PR, 1).forEach((x, i) => { if (norm_(x[0]) === norm_(product)) sh_(PR).getRange(i + 2, 7).setValue(unit); });
    return { amount: amt, pack: pack, received: rec, unit: unit, planned: +g.x[4], kind: kind };
  });
}
function batchDetail_(id) {
  return { issues: rows_(BX, 6).filter(r => String(r[0]) === String(id)).map(r => ({ date: fd(r[1]), item: r[2], qty: r[3], rate: r[4], amt: r[5] })),
    receipts: rows_(SIN, 7).filter(r => r[6] === 'BATCH' && String(r[5]) === String(id)).map(r => ({ date: fd(r[0]), qty: r[2] })) };
}

/* ---------- Excel download ---------- */
function exportXlsx_(from, to) {
  const f = from || '0000-00-00', t = to || '9999-99-99', inR = d => { const k = fd(d); return k >= f && k <= t; };
  const ov = rows_(DATA, 17).filter(x => x[0] && inR(x[1])), iv = rows_(IT, 10).filter(x => x[0] && inR(x[1])), st = {};
  rows_(DATA, 17).forEach(x => st[+x[0]] = x[6]);
  const tmp = SpreadsheetApp.create('tmp_export');
  try {
    const s1 = tmp.getSheets()[0].setName('Sheet1'), h1 = ['DATE','ORDER NO','ORDER TAKER','CUSTOMER','NUMBER','CITY','ADDRESS','ITEMS','AMOUNT','D.C','TOTAL','ADVANCE','COD','TOTAL QTY','STATUS','TRACKING ID','DELIVERY'];
    s1.getRange(1, 1, 1, h1.length).setValues([h1]).setFontWeight('bold');
    if (ov.length) {
      s1.getRange(2, 5, ov.length, 1).setNumberFormat('@');
      s1.getRange(2, 1, ov.length, 17).setValues(ov.map(x => [x[1], x[0], x[10], x[2], x[3], x[4], x[9], x[11], x[12], x[13], '', x[14], '', x[15], x[6], x[7], x[16]]));
      s1.getRange(2, 11, ov.length, 1).setFormulas(ov.map((x, i) => ['=I' + (i + 2) + '+J' + (i + 2)]));
      s1.getRange(2, 13, ov.length, 1).setFormulas(ov.map((x, i) => ['=K' + (i + 2) + '-L' + (i + 2)]));
      s1.getRange(2, 1, ov.length, 1).setNumberFormat('yyyy-mm-dd');
    }
    const s2 = tmp.insertSheet('All Orders Data'), h2 = ['Order #','Date','Month','Year','Product','Qty','Unit Price (Rs)','Amount (Rs)','STATUS','TOTAL UNITS'];
    s2.getRange(1, 1, 1, 10).setValues([h2]).setFontWeight('bold');
    if (iv.length) { s2.getRange(2, 1, iv.length, 10).setValues(iv.map(x => [x[0], x[1], x[2], x[3], x[4], x[5], x[6], x[7], st[+x[0]] || '', x[9]])); s2.getRange(2, 2, iv.length, 1).setNumberFormat('yyyy-mm-dd'); }
    SpreadsheetApp.flush();
    const b = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    return { name: 'ORDERS_' + (from || 'all') + '_' + (to || 'all') + '.xlsx', b64: Utilities.base64Encode(b.getBytes()), orders: ov.length };
  } finally { DriveApp.getFileById(tmp.getId()).setTrashed(true); }
}
