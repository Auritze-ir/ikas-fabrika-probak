/**
 * IKAS FABRIKA · Apps Script
 * ------------------------------------------------------------
 * ERABILERA (","):
 *  1. script.google.com -> Proiektu berria. Itsatsi fitxategi hau (Code.gs).
 *  2. Gehitu bi HTML fitxategi (Fitxategia > Gehitu > HTML):
 *       "Salmenta"  -> Ikas_fabrika_salmenta.html edukia
 *       "Barne"     -> Ikas_fabrika_barne.html edukia
 *  3. Exekutatu  setup()  (baimenak onartu). Honek:
 *       - "Ikas fabrika" Google Sheet-a sortzen du (Eskaerak, Stock, Arduradunak, Konfigurazioa)
 *       - barne-orrirako gako sekretua sortzen du
 *       - egunero stock-a eta atzerapenak begiratzeko tarte-aktibatzailea (trigger) jartzen du
 *  4. Sheet-eko "Arduradunak" orrian jarri ikasle arduradunen posta-helbideak.
 *  5. Zabaldu > Inplementazio berria > Web aplikazioa:
 *       Exekutatu: Ni  |  Sarbidea: Edonork
 *     Salmenta-orria: web aplikazioaren helbidea
 *     Barne-orria:    helbidea + ?orria=barne&gakoa=GAKOA   (setup() eta emailek erakusten dute)
 *  6. Kodea aldatzen baduzu: Inplementazioa kudeatu > Editatu > Bertsio berria.
 * ------------------------------------------------------------ */

const COLS = ['id','createdAt','ref','name','contact','group','product','use','size','qty','due','notes','source','status','tasks','meeting','decisions'];
const STOCK_COLS = ['product','size','label','g','stock','min','refill'];
const STATUSES = ['jaso','bilera','plan','prest','probak','entregatua'];
const DEFAULTS = { LEAD_DAYS:7, CURING_DAYS:90 };
const PRODUCT_NAMES = { solidoa:'Xaboi solidoa', hautsa:'Xaboi hautsa' };

/* ---------- Konfigurazioa eta Sheet-a ---------- */
function props_(){ return PropertiesService.getScriptProperties(); }
function ss_(){
  const id = props_().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) { props_().setProperty('SS_ID', active.getId()); return active; }
  throw new Error('Exekutatu lehenengo setup()');
}
function sheet_(name){ return ss_().getSheetByName(name); }
function conf_(){
  const out = Object.assign({}, DEFAULTS), sh = sheet_('Konfigurazioa');
  if (sh && sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r => { if (r[0]) out[r[0]] = r[1]; });
  out.LEAD_DAYS = Number(out.LEAD_DAYS) || DEFAULTS.LEAD_DAYS;
  out.CURING_DAYS = Number(out.CURING_DAYS) || DEFAULTS.CURING_DAYS;
  return out;
}
function today_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function addDays_(n){ const d = new Date(); d.setDate(d.getDate()+n); return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function dateStr_(v){ return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v||''); }
function checkKey_(key){
  const real = props_().getProperty('ACCESS_KEY');
  if (!real || key !== real) throw new Error('Baimenik ez');
}

/* ---------- SETUP: Sheet-a eta aktibatzaileak sortu ---------- */
function setup(){
  let ss;
  const existing = props_().getProperty('SS_ID');
  if (existing) { ss = SpreadsheetApp.openById(existing); }
  else {
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Ikas fabrika');
    props_().setProperty('SS_ID', ss.getId());
  }
  const ensure = (name, header, widths) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#e4f1ec'); sh.setFrozenRows(1); }
    return sh;
  };
  const eskaerak = ensure('Eskaerak', COLS);
  eskaerak.getRange('C:C').setNumberFormat('@'); eskaerak.getRange('K:K').setNumberFormat('@');
  eskaerak.getRange('N2:N').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());

  const stock = ensure('Stock', STOCK_COLS);
  if (stock.getLastRow() < 2) stock.getRange(2,1,4,7).setValues([
    ['solidoa','txikia','Xaboi solidoa · Txikia',50,60,20,80],
    ['solidoa','ertaina','Xaboi solidoa · Ertaina',100,60,20,80],
    ['solidoa','handia','Xaboi solidoa · Handia',200,30,10,40],
    ['hautsa','poltsa','Xaboi hautsa · Poltsa',100,40,15,60]]);

  const arduradunak = ensure('Arduradunak', ['arloa','izena','email']);
  if (arduradunak.getLastRow() < 2) arduradunak.getRange(2,1,1,3).setValues([['Kimika','(jarri izena)', Session.getEffectiveUser().getEmail()]]);

  const konf = ensure('Konfigurazioa', ['gakoa','balioa','azalpena']);
  if (konf.getLastRow() < 2) konf.getRange(2,1,2,3).setValues([
    ['LEAD_DAYS', DEFAULTS.LEAD_DAYS, 'Stock-etik entregatzeko gutxieneko egunak'],
    ['CURING_DAYS', DEFAULTS.CURING_DAYS, 'Stock berritzeko behar den denbora (kurutzea, egunak)']]);

  if (!props_().getProperty('ACCESS_KEY')) props_().setProperty('ACCESS_KEY', Utilities.getUuid().replace(/-/g,''));

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyCheck').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyCheck').timeBased().everyDays(1).atHour(7).create();

  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Hoja 1') || ss.getSheetByName('Hoja de cálculo 1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  Logger.log('Sheet-a: ' + ss.getUrl());
  Logger.log('Barne-orriko gakoa: ' + props_().getProperty('ACCESS_KEY'));
  Logger.log('Betetzeko: "Arduradunak" orrian jarri ikasleen posta-helbideak.');
}
function baseUrl_(){ try { return ScriptApp.getService().getUrl(); } catch(e){ return ''; } }

/* ---------- Web orriak ---------- */
function doGet(e){
  const page = (e && e.parameter && e.parameter.orria) || '';
  if (page === 'barne') {
    const key = (e.parameter.gakoa || '');
    if (key !== props_().getProperty('ACCESS_KEY')) return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">Baimenik ez. Erabili arduradunen helbide osoa.</p>');
    const html = HtmlService.createHtmlOutputFromFile('Barne').getContent().split('%%KEY%%').join(key).split('%%SHOP_URL%%').join(baseUrl_());
    return HtmlService.createHtmlOutput(html).setTitle('Ikas fabrika · barne-orria').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Salmenta').setTitle('Ikas fabrika · xaboiak')
    .addMetaTag('viewport','width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Stock ---------- */
function stockRows_(){
  const sh = sheet_('Stock'); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,STOCK_COLS.length).getValues().map((r,i) => {
    const o = {}; STOCK_COLS.forEach((c,j) => o[c] = r[j]); o._row = i+2; return o; });
}
/** Publikoa: salmenta-orriak erabiltzen du. */
function getStock(){
  return { leadDays: conf_().LEAD_DAYS, stock: stockRows_().map(r => ({ product:r.product, size:r.size, stock:Number(r.stock)||0 })) };
}
/** Barnekoa: gakoarekin. */
function getStockAdmin(key){
  checkKey_(key);
  return stockRows_().map(r => ({ product:r.product, size:r.size, stock:Number(r.stock)||0, min:Number(r.min)||0, refill:Number(r.refill)||1 }));
}
function setStock(key, r){
  checkKey_(key);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const row = stockRows_().find(x => x.product === r.product && x.size === r.size);
    if (!row) throw new Error('Produktua ez da aurkitu');
    sheet_('Stock').getRange(row._row,5,1,3).setValues([[Number(r.stock)||0, Number(r.min)||0, Math.max(1,Number(r.refill)||1)]]);
    refillIfNeeded_([{product:r.product,size:r.size}]);
    return true;
  } finally { lock.releaseLock(); }
}

/* ---------- Eskaerak ---------- */
function nextId_(){
  const n = (Number(props_().getProperty('SEQ')) || 0) + 1; props_().setProperty('SEQ', String(n));
  return 'ESK-' + ('000' + n).slice(-3);
}
function readOrders_(){
  const sh = sheet_('Eskaerak'); if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,COLS.length).getValues().map((r,i) => {
    const o = {}; COLS.forEach((c,j) => o[c] = r[j]);
    o.due = dateStr_(o.due); o.createdAt = r[1] instanceof Date ? r[1].toISOString() : String(r[1]||''); o._row = i+2; return o; });
}
function appendOrder_(o){
  const row = COLS.map(c => o[c] == null ? '' : o[c]);
  sheet_('Eskaerak').appendRow(row);
}
/** Barnekoa */
function getOrders(key){
  checkKey_(key);
  return readOrders_().map(o => { delete o._row; return o; });
}
/** Barnekoa: txartela eguneratu. "stock" motako eskaera entregatzen denean stock-a handitzen da. */
function updateOrder(key, o){
  checkKey_(key);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const old = readOrders_().find(x => x.id === o.id);
    if (!old) throw new Error('Eskaera ez da aurkitu');
    if (old.source === 'stock' && old.status !== 'entregatua' && o.status === 'entregatua') {
      const r = stockRows_().find(x => x.product === old.product && x.size === old.size);
      if (r) sheet_('Stock').getRange(r._row,5).setValue((Number(r.stock)||0) + (Number(old.qty)||0));
      notify_('Stock-a berritu da', [PRODUCT_NAMES[old.product] + ' · ' + old.size + ': +' + old.qty + ' ale stock-ean.']);
    }
    const row = COLS.map(c => (c in o) ? o[c] : old[c]);
    // aldaezinak mantendu
    ['id','createdAt','ref','source'].forEach(c => row[COLS.indexOf(c)] = old[c]);
    sheet_('Eskaerak').getRange(old._row,1,1,COLS.length).setValues([row]);
    return true;
  } finally { lock.releaseLock(); }
}
/** Publikoa: salmenta-orritik. Stock-a egiaztatu, murriztu eta ikasleei abisatu. */
function addOrder(order){
  const conf = conf_();
  const name = String(order.name||'').trim().slice(0,120), contact = String(order.contact||'').trim().slice(0,160);
  if (!name || !contact) throw new Error('Datuak falta dira');
  const lines = (order.lines||[]).map(l => ({ product:String(l.product), size:String(l.size), use:String(l.use||''), qty:Math.floor(Number(l.qty)) }))
    .filter(l => l.qty > 0);
  if (!lines.length || lines.length > 20) throw new Error('Eskaera hutsik');
  const due = String(order.due||'');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(due) || due < addDays_(conf.LEAD_DAYS)) throw new Error('Data ez da egokia');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const stock = stockRows_(), need = {};
    lines.forEach(l => { const k = l.product+'|'+l.size; need[k] = (need[k]||0) + l.qty; });
    Object.keys(need).forEach(k => { const r = stock.find(x => x.product+'|'+x.size === k);
      if (!r || (Number(r.stock)||0) < need[k]) throw new Error('Stock nahikorik ez: ' + k); });
    Object.keys(need).forEach(k => { const r = stock.find(x => x.product+'|'+x.size === k);
      sheet_('Stock').getRange(r._row,5).setValue((Number(r.stock)||0) - need[k]); });
    const ids = [], ref = nextId_();
    lines.forEach((l,i) => {
      const id = i === 0 ? ref : nextId_(); ids.push(id);
      appendOrder_({ id:id, createdAt:new Date(), ref:ref, name:name, contact:contact, group:String(order.group||'').slice(0,120),
        product:l.product, use:l.use, size:l.size, qty:l.qty, due:due, notes:String(order.notes||'').slice(0,1000),
        source:'web', status:'jaso', tasks:'{}', meeting:'{}', decisions:'[]' });
    });
    const refills = refillIfNeeded_(Object.keys(need).map(k => ({product:k.split('|')[0], size:k.split('|')[1]})));
    const summary = lines.map(l => '• ' + (PRODUCT_NAMES[l.product]||l.product) + ' · ' + l.size + ' · ' + l.qty + ' ale · ' + l.use);
    notify_('Eskaera berria: ' + ref + ' (' + name + ')', ['Eskaera jaso da salmenta-orritik.', ''].concat(summary, ['', 'Izena: ' + name, 'Harremana: ' + contact, 'Entrega-data: ' + due, order.notes ? 'Oharrak: ' + order.notes : '']
      .filter(x => x !== ''), refills.length ? ['', 'Stock berritzea behar da: ' + refills.join(', ')] : []));
    return { id: ref, ids: ids };
  } finally { lock.releaseLock(); }
}

/* ---------- Stock berritzea (automatikoa) ---------- */
/** Stock <= minimoa bada eta irekita berritze-eskaerarik ez badago, sortu bat. Deitu lock barruan. */
function refillIfNeeded_(pairs){
  const conf = conf_(), made = [];
  const orders = readOrders_(), stock = stockRows_();
  pairs.forEach(p => {
    const r = stock.find(x => x.product === p.product && x.size === p.size); if (!r) return;
    const low = (Number(r.stock)||0) <= (Number(r.min)||0);
    const open = orders.some(o => o.source === 'stock' && o.product === p.product && o.size === p.size && o.status !== 'entregatua');
    if (low && !open) {
      const id = nextId_();
      appendOrder_({ id:id, createdAt:new Date(), ref:id, name:'Stock berritzea', contact:'', group:'Kimika', product:p.product, use:'', size:p.size,
        qty:Number(r.refill)||1, due:addDays_(conf.CURING_DAYS), notes:'Stock-a minimoaren azpitik dago (' + r.stock + '/' + r.min + '). Lote berria egin behar da; kurutzeak ' + conf.CURING_DAYS + ' egun behar ditu.',
        source:'stock', status:'jaso', tasks:'{}', meeting:'{}', decisions:'[]' });
      orders.push({source:'stock',product:p.product,size:p.size,status:'jaso'});
      made.push((PRODUCT_NAMES[p.product]||p.product) + ' ' + p.size + ' (' + id + ')');
    }
  });
  return made;
}
/** Egunero 07:00: stock baxua berrikusi eta atzeratutako eskaerak abisatu. */
function dailyCheck(){
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let made = [], late = [];
  try {
    made = refillIfNeeded_(stockRows_().map(r => ({product:r.product,size:r.size})));
    const t = today_();
    late = readOrders_().filter(o => o.status !== 'entregatua' && o.due && o.due < t).map(o => o.id + ' · ' + (o.name||'') + ' · ' + o.due);
  } finally { lock.releaseLock(); }
  const lines = [];
  if (made.length) lines.push('Stock berritze-eskaera berriak: ' + made.join(', '));
  if (late.length) lines.push('Atzeratutako eskaerak:'), late.forEach(x => lines.push('• ' + x));
  if (lines.length) notify_('Ikas fabrika · eguneroko laburpena', lines);
}

/* ---------- Posta (agentea) ---------- */
function notify_(subject, lines){
  const sh = sheet_('Arduradunak'); if (!sh || sh.getLastRow() < 2) return;
  const to = sh.getRange(2,3,sh.getLastRow()-1,1).getValues().map(r => String(r[0]).trim()).filter(x => /@/.test(x));
  if (!to.length) return;
  const key = props_().getProperty('ACCESS_KEY'), link = baseUrl_() + '?orria=barne&gakoa=' + key;
  const body = lines.join('\n') + '\n\nEskaera-taula: ' + link + '\n\nIkas fabrika';
  try { MailApp.sendEmail({ to: to.join(','), subject: '[Ikas fabrika] ' + subject, body: body, name: 'Ikas fabrika' }); }
  catch (e) { console.error('Posta ezin izan da bidali: ' + e); }
}
