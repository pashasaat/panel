/**
 * ERSA · Saat & Gözlük — Apps Script backend (mağaza + yönetici panel)
 * Kurulum sonrası DEĞİŞİKLİK yaptıysan: Dağıt → Dağıtımları yönet → ✏️ →
 * Sürüm: "Yeni sürüm" → Dağıt  (URL aynı kalır)
 *
 * PANEL İÇİN: Ayarlar sayfasına bir satır ekle →  adminPin | <gizli-pin>
 * Hiçbir veri SİLİNMEZ; her şey eklenir/güncellenir.
 */

var SHEETS = { products: 'Urunler', orders: 'Siparisler', offers: 'Teklifler', settings: 'Ayarlar' };
var PUBLIC_SETTINGS = ['usdTry', 'brand', 'whatsapp', 'freeShipping'];

/* ================= GET ================= */
function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var action = p.action || 'products';

    if (action === 'products') {
      return json_({ ok: true, products: getProducts_(false), settings: getPublicSettings_() });
    }
    if (action === 'admin') {
      var g = checkPin_(p.pin);
      if (g !== 'OK') return json_({ ok: false, error: g === 'NO_PIN' ? 'Ayarlar sayfasına adminPin ekleyin' : 'PIN hatalı' });
      return json_({
        ok: true,
        products: getSheetObjects_(SHEETS.products),
        orders: getSheetObjects_(SHEETS.orders),
        offers: getSheetObjects_(SHEETS.offers),
        settings: getPublicSettings_()
      });
    }
    return json_({ ok: false, error: 'bilinmeyen action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ================= POST ================= */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // herkese açık (mağaza)
    if (d.action === 'order') return addOrder_(d);
    if (d.action === 'offer') return addOffer_(d);

    // PIN korumalı (panel)
    var admin = { save_product: 1, set_status: 1, save_setting: 1 };
    if (admin[d.action]) {
      var g = checkPin_(d.pin);
      if (g !== 'OK') return json_({ ok: false, error: g === 'NO_PIN' ? 'adminPin tanımsız' : 'PIN hatalı' });
      if (d.action === 'save_product') return saveProduct_(d);
      if (d.action === 'set_status') return setStatus_(d);
      if (d.action === 'save_setting') return saveSetting_(d);
    }
    return json_({ ok: false, error: 'bilinmeyen action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ============ MAĞAZA: ürünler ============ */
function getProducts_(includeInactive) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.products);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row[0]) continue;
    var p = { specs: {} };
    for (var c = 0; c < headers.length; c++) {
      var rawKey = String(headers[c]).trim(), key = rawKey.toLowerCase(), val = row[c];
      if (key === 'id') p.id = String(val).trim();
      else if (key === 'type') p.type = String(val).trim().toLowerCase();
      else if (key === 'brand') p.brand = String(val).trim();
      else if (key === 'model') p.model = String(val).trim();
      else if (key === 'name') p.name = String(val).trim();
      else if (key === 'priceusd') p.priceUSD = Number(val) || 0;
      else if (key === 'condition') p.condition = String(val).trim();
      else if (key === 'listedat') p.listedAt = fmtDate_(val);
      else if (key === 'qty') p.qty = Number(val) || 1;
      else if (key === 'active') p.active = isTrue_(val);
      else if (key === 'images') p.images = splitImages_(val);
      else if (rawKey && val !== '' && val != null) p.specs[rawKey] = String(val).trim();
    }
    if (!includeInactive && p.active === false) continue;
    if (!p.type) p.type = 'saat';
    if (!p.images) p.images = [];
    out.push(p);
  }
  return out;
}

/* ============ PANEL: ham satır okuma ============ */
function getSheetObjects_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return { headers: [], rows: [] };
  var values = sh.getDataRange().getValues();
  if (values.length < 1) return { headers: [], rows: [] };
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var obj = { _row: r + 1 }, any = false;
    for (var c = 0; c < headers.length; c++) {
      var v = values[r][c];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      obj[headers[c]] = v;
      if (v !== '' && v != null) any = true;
    }
    if (any) rows.push(obj);
  }
  return { headers: headers, rows: rows };
}

/* ============ PANEL: ürün kaydet (upsert) ============ */
function saveProduct_(d) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.products) || ss.insertSheet(SHEETS.products);
  var values = sh.getDataRange().getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idCol = headers.indexOf('id');
  if (idCol < 0) return json_({ ok: false, error: 'id kolonu yok' });
  var row = d.row || {};
  var id = String(row.id || '').trim();
  if (!id) return json_({ ok: false, error: 'id boş olamaz' });

  var target = -1;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol]).trim() === id) { target = r + 1; break; }
  }
  if (target > 0) {
    for (var c = 0; c < headers.length; c++) {
      if (headers[c] in row) sh.getRange(target, c + 1).setValue(row[headers[c]]);
    }
    return json_({ ok: true, id: id, mode: 'update' });
  }
  var arr = headers.map(function (h) { return (h in row) ? row[h] : ''; });
  sh.appendRow(arr);
  return json_({ ok: true, id: id, mode: 'insert' });
}

/* ============ PANEL: sipariş/teklif durumu ============ */
function setStatus_(d) {
  var name = d.sheet;
  if ([SHEETS.orders, SHEETS.offers].indexOf(name) < 0) return json_({ ok: false, error: 'geçersiz sheet' });
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return json_({ ok: false, error: 'sheet yok' });
  var headers = sh.getDataRange().getValues()[0].map(function (h) { return String(h).trim(); });
  var col = headers.indexOf('Durum');
  if (col < 0) return json_({ ok: false, error: 'Durum kolonu yok' });
  var r = Number(d.row);
  if (!r || r < 2) return json_({ ok: false, error: 'geçersiz satır' });
  sh.getRange(r, col + 1).setValue(d.status || '');
  return json_({ ok: true });
}

/* ============ PANEL: ayar kaydet ============ */
function saveSetting_(d) {
  var sh = getOrCreate_(SHEETS.settings, ['anahtar', 'deger', 'aciklama']);
  var values = sh.getDataRange().getValues();
  var k = String(d.key || '').trim();
  if (!k) return json_({ ok: false, error: 'anahtar boş' });
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === k) { sh.getRange(r + 1, 2).setValue(d.value); return json_({ ok: true }); }
  }
  sh.appendRow([k, d.value, '']);
  return json_({ ok: true });
}

/* ============ MAĞAZA: sipariş/teklif ekle ============ */
function addOrder_(d) {
  var sh = getOrCreate_(SHEETS.orders,
    ['Tarih', 'Durum', 'UrunIDler', 'Urunler', 'TutarUSD', 'ParaBirimi', 'Sigorta', 'Musteri', 'Telefon', 'Adres', 'Not']);
  var items = d.items || [];
  var total = items.reduce(function (s, i) { return s + (Number(i.priceUSD) || 0); }, 0) + (d.insurance ? 5 : 0);
  sh.appendRow([new Date(), 'Yeni',
    items.map(function (i) { return i.id; }).join(', '),
    items.map(function (i) { return i.name; }).join(' | '),
    total, d.currency || '', d.insurance ? 'Evet' : 'Hayır',
    d.customer || '', d.phone || '', d.address || '', d.note || '']);
  return json_({ ok: true });
}
function addOffer_(d) {
  var sh = getOrCreate_(SHEETS.offers,
    ['Tarih', 'Durum', 'UrunID', 'Urun', 'Turler', 'NakitTeklif', 'TakasUrun', 'Not', 'Musteri', 'Telefon']);
  var name = '';
  try { var ps = getProducts_(true); for (var i = 0; i < ps.length; i++) if (ps[i].id === d.productId) name = ps[i].name; } catch (e) {}
  sh.appendRow([new Date(), 'Yeni', d.productId || '', name, (d.types || []).join(', '),
    d.cash || '', d.trade || '', d.note || '', d.customer || '', d.phone || '']);
  return json_({ ok: true });
}

/* ============ AYARLAR ============ */
function getSettings_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.settings);
  var s = {};
  if (!sh) return s;
  var v = sh.getDataRange().getValues();
  for (var r = 1; r < v.length; r++) { var k = String(v[r][0]).trim(); if (k) s[k] = v[r][1]; }
  return s;
}
function getPublicSettings_() {
  var all = getSettings_(), out = {};
  PUBLIC_SETTINGS.forEach(function (k) { if (k in all) out[k] = all[k]; });
  return out;
}
function checkPin_(pin) {
  var real = String(getSettings_().adminPin || '').trim();
  if (!real) return 'NO_PIN';
  return String(pin || '').trim() === real ? 'OK' : 'BAD';
}

/* ============ yardımcı ============ */
function splitImages_(val) {
  if (!val) return [];
  return String(val).split(/[\n,;]+/).map(function (s) { return s.trim(); }).filter(String);
}
function isTrue_(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return !(s === 'false' || s === '0' || s === 'hayır' || s === 'hayir' || s === 'no' || s === '');
}
function fmtDate_(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  return String(v).trim();
}
function getOrCreate_(name, headers) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); }
  return sh;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
