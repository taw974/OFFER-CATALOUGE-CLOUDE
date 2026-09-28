/**
 * OFFER ENTRY — the master sheet behind the flyer            (version 2.4)
 * ------------------------------------------------------------------------
 * This file is the MASTER: it holds the DATABASE and runs everything.
 * Branch sheets are separate files made from here (Offer Tools → Create a
 * branch sheet). A barcode typed in any branch sheet is looked up in this
 * master's database; a product typed by hand is saved into it, so every
 * branch learns it. A new PRODUCT MASTER is imported once, here.
 *
 * Do not "Make a copy" of a branch sheet or of this file for a branch —
 * a copy has no trigger. Use Create a branch sheet instead.
 *
 * Set up once (the owner of the file): menu  Offer Tools → 1. Set up
 *
 * 2.2: rows can be inserted and deleted in the offer tabs (also in branch
 * sheets). Only the heading row is locked; SL and BRAND → ROW STATE show a
 * warning when typed in. After pasting this version, run 1. Set up again once —
 * it unlocks the branch sheets that already exist.
 *
 * 2.3: CATEGORY comes after OFFER PRICE (1. Set up moves it in every tab; a tab
 * not moved yet moves itself on its first edit). A barcode fills faster: the tab
 * is read once, fewer writes, and branches no longer wait for each other.
 *
 * 2.4: columns are SL, PRODUCT NAME, BARCODE, C.P, S.P, OFFER PRICE, NOTE,
 * CATEGORY, ... (1. Set up puts every tab in this order; any tab still in an
 * older order is put right on its first edit). The DATABASE is kept in the
 * script cache, so a scanned barcode no longer opens the master file.
 */

var CFG = {
  HEAD: ['SL','PRODUCT NAME','BARCODE','C.P','S.P','OFFER PRICE','NOTE','CATEGORY','BRAND','PACKING','CHECK','ROW STATE'],
  C: {SL:1, NAME:2, BC:3, CP:4, SP:5, OFFER:6, NOTE:7, CAT:8, BRAND:9, PACK:10, CHECK:11, STATE:12},
  FIRST: 2, ROWS: 500,                 /* ROWS: the rows a new tab starts with; the tab may grow or shrink */
  DB: 'DATABASE', MAP: 'CATEGORY MAP', LOG: 'UPDATES', GUIDE: 'GUIDE', BR: 'BRANCH SHEETS',
  BRH: ['BRANCH', 'LINK', 'SPREADSHEET ID', 'CREATED'],
  DBH: ['BARCODE','KEY','PRODUCT NAME','BRAND','ITEM CODE','PACKING','UNIT','FAMILY','GROUP','CATEGORY','NAME SET BY','CATEGORY SET BY','SOURCE','UPDATED'],
  D: {BC:1, KEY:2, NAME:3, BRAND:4, CODE:5, PACK:6, UNIT:7, FAM:8, GRP:9, CAT:10, NAMEBY:11, CATBY:12, SRC:13, UPD:14},
  CATS: ['Fresh / Vegetable','Fresh Fish','Meat','Chicken','Rice','Oil','Grocery','Household','Cosmetics','Mobile & Electronics','Garments','Other'],
  BULK: 12
};

/* ============================== menu ============================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Offer Tools')
    .addItem('1. Set up (first time only)', 'setup')
    .addSeparator()
    .addItem('Create a branch sheet', 'createBranchSheet')
    .addItem('Show the branch sheets', 'showBranchSheets')
    .addSeparator()
    .addItem('Update the database from an imported PRODUCT MASTER', 'updateFromMaster')
    .addItem('Fill every branch sheet again from the database', 'refillAllBranches')
    .addItem('Apply the CATEGORY MAP to the database again', 'reapplyCategories')
    .addSeparator()
    .addItem('Check this tab again', 'recheckActiveTab')
    .addItem('Refresh the barcode cache', 'warmCacheMenu')
    .addToUi();
}
/* ======================= the master and its branch sheets ======================= */
var FP_MASTER = null;
/** the master spreadsheet — the one that holds the DATABASE — whichever sheet was edited */
function fpMaster() {
  if (FP_MASTER) return FP_MASTER;
  var id = '';
  try { id = PropertiesService.getScriptProperties().getProperty('MASTER_ID') || ''; } catch (err) {}
  if (id) { try { FP_MASTER = SpreadsheetApp.openById(id); return FP_MASTER; } catch (err2) {} }
  FP_MASTER = SpreadsheetApp.getActive();
  return FP_MASTER;
}
function fpBranchList() {
  var br = fpMaster().getSheetByName(CFG.BR); if (!br) return [];
  var n = Math.max(0, br.getLastRow() - 1); if (!n) return [];
  return br.getRange(2, 1, n, CFG.BRH.length).getValues()
    .filter(function (r) { return String(r[2] || '').trim(); })
    .map(function (r) { return {name: String(r[0]), link: String(r[1]), id: String(r[2]).trim()}; });
}
function fpHasTriggerFor(id, fn) {
  fn = fn || 'handleEdit';
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === fn && String(t.getTriggerSourceId && t.getTriggerSourceId()) === String(id);
  });
}
/** handleEdit for typing, handleChange for rows inserted or deleted */
function fpEnsureTrigger(ssOrId) {
  var id = typeof ssOrId === 'string' ? ssOrId : ssOrId.getId();
  if (!fpHasTriggerFor(id, 'handleEdit')) ScriptApp.newTrigger('handleEdit').forSpreadsheet(id).onEdit().create();
  if (!fpHasTriggerFor(id, 'handleChange')) ScriptApp.newTrigger('handleChange').forSpreadsheet(id).onChange().create();
}
function fpIsLinked(id) { return fpHasTriggerFor(id, 'handleEdit') && fpHasTriggerFor(id, 'handleChange'); }
function fpEnsureBranchTab(ss) {
  var br = ss.getSheetByName(CFG.BR);
  if (!br) {
    br = ss.insertSheet(CFG.BR);
    br.getRange(1, 1, 1, CFG.BRH.length).setValues([CFG.BRH]);
  }
  return br;
}

/* ======================= pure helpers (no Google calls) ======================= */
function fpText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return isFinite(v) ? v.toFixed(0) : '';
  return String(v).trim();
}
/** the barcode without spaces and without leading zeros — only a fallback for pasted numbers */
function fpKey(v) {
  var s = fpText(v).replace(/\s+/g, '').toUpperCase();
  return /^\d+$/.test(s) ? s.replace(/^0+/, '') : s;
}
function fpIsSinglePack(p) { return /^(PCS?|EA|EACH|NOS?|KGS?|G|GM|L|LTR|ML|)$/i.test(String(p || '').trim()); }
/** a barcode on two lines (single and carton): the single one goes on offer */
function fpPick(hits) {
  if (!hits || !hits.length) return null;
  for (var i = 0; i < hits.length; i++) if (fpIsSinglePack(hits[i].vals[CFG.D.PACK - 1])) return hits[i];
  return hits[0];
}
function fpNum(v) {
  if (v === '' || v === null || v === undefined) return '';
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isFinite(n) ? String(Math.round(n * 1000) / 1000) : String(v).trim().toUpperCase();
}
function fpStateLabel(st) {
  if (st === 'db') return '✓ In database';
  if (st === 'new') return '★ New — type the name';
  if (st === 'added') return '★ New — saved to database';
  if (st === 'nobc') return '• No barcode';
  return '';
}
/** the same line twice = same barcode, name and all three prices. Same name, other price: fine. */
function fpDuplicates(rows) {           // rows: [barcode, name, category, cp, sp, offer]
  var seen = {}, out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], bc = fpText(r[0]).toUpperCase(), nm = String(r[1] || '').trim().toUpperCase();
    if (!bc && !nm) { out.push(0); continue; }
    var k = [bc, nm, fpNum(r[3]), fpNum(r[4]), fpNum(r[5])].join('|');
    if (seen[k]) out.push(seen[k]); else { seen[k] = i + 1; out.push(0); }
  }
  return out;                            // for each row: the SL of the first identical row, or 0
}
/** a category for a family/group the map has never seen */
function fpGuessCat(fam, grp) {
  var F = String(fam || '').toLowerCase(), G = String(grp || '').toLowerCase();
  if (/fruit|vegetable/.test(F)) return 'Fresh / Vegetable';
  if (/fish/.test(F + ' ' + G)) return 'Fresh Fish';
  if (/chicken|poultry|egg/.test(F + ' ' + G)) return 'Chicken';
  if (/meat|burger/.test(F + ' ' + G)) return 'Meat';
  if (/^rice$/.test(G)) return 'Rice';
  if (/oil|ghee/.test(G)) return 'Oil';
  if (/health|beauty|cosmetic|perfume|female|medicated/.test(F + ' ' + G)) return 'Cosmetics';
  if (/electronic|it products|appliance|mobile|watch|recharge|computer/.test(F + ' ' + G)) return 'Mobile & Electronics';
  if (/garment|wear|footwear|luggage|bag|jewel|kids/.test(F + ' ' + G)) return 'Garments';
  if (/household|non foods|furnishing|hardware|kitchen|clean|laundry/.test(F + ' ' + G)) return 'Household';
  if (/grocery|dairy|frozen|bakery|roastery|delicatessen|food/.test(F)) return 'Grocery';
  return 'Other';
}
/** the POS report "Item Master By Group": a heading row, then items, with family/group lines in between */
function fpParseMaster(grid) {
  var hi = -1, col = {};
  for (var r = 0; r < Math.min(grid.length, 80); r++) {
    var low = grid[r].map(function (v) { return String(v === null || v === undefined ? '' : v).trim().toLowerCase(); });
    if (low.indexOf('barcode') >= 0 && low.indexOf('description') >= 0) {
      hi = r; low.forEach(function (v, i) { if (v && col[v] === undefined) col[v] = i; }); break;
    }
  }
  if (hi < 0) throw new Error('This tab has no "Barcode" and "Description" heading — is it the PRODUCT MASTER?');
  var pick = function (names) { for (var i = 0; i < names.length; i++) if (col[names[i]] !== undefined) return col[names[i]]; return -1; };
  var cBc = col['barcode'], cDesc = col['description'], cBrand = pick(['brand']), cCode = pick(['item code', 'code']),
      cPack = pick(['packing', 'pack']), cUm = pick(['u.m.', 'u.m', 'um', 'unit']);
  var out = [], fam = '', grp = '', lastWasHead = false, lastHead = '';
  for (var r2 = hi + 1; r2 < grid.length; r2++) {
    var row = grid[r2];
    var get = function (c) { return c < 0 ? '' : String(row[c] === null || row[c] === undefined ? '' : row[c]).trim(); };
    var bc = fpText(row[cBc]), desc = get(cDesc);
    if (bc && desc && bc.toLowerCase() !== 'barcode') {
      var brand = get(cBrand); if (/^none$/i.test(brand)) brand = '';
      out.push({bc: bc, name: desc, brand: brand, code: fpText(row[cCode]), pack: get(cPack), unit: get(cUm), fam: fam, grp: grp});
      lastWasHead = false; continue;
    }
    var filled = [];
    for (var c = 0; c < row.length; c++) if (String(row[c] === null || row[c] === undefined ? '' : row[c]).trim() !== '') filled.push(c);
    if (!filled.length) continue;                                   // a blank line changes nothing
    var only = filled.length === 1 && filled[0] < cBc;
    var text = only ? String(row[filled[0]]).trim() : '';
    if (only && text && !/:\s*$/.test(text) && !/^page\s+\d+/i.test(text)) {
      if (lastWasHead) { fam = lastHead; grp = text; }              // two headings in a row: family, then group
      else grp = text;
      lastWasHead = true; lastHead = text; continue;
    }
    lastWasHead = false;                                           // a page header or anything else
  }
  return out;
}
/** merge a parsed master into the database rows (arrays of 14). Returns counts. */
function fpMerge(rows, items, catMap, today) {
  var D = CFG.D, idx = {}, byBc = {};
  rows.forEach(function (r, i) {
    idx[fpText(r[D.BC - 1]).toUpperCase() + '|' + String(r[D.PACK - 1] || '').toUpperCase() + '|' + String(r[D.NAME - 1] || '').toUpperCase()] = i;
    var b = fpText(r[D.BC - 1]).toUpperCase(); (byBc[b] = byBc[b] || []).push(i);
  });
  var added = 0, changed = 0, same = 0, touched = {}, seen = {};
  items.forEach(function (it) {
    var full = [it.bc, it.name, it.brand, it.code, it.pack, it.unit, it.fam, it.grp].join('\u0001');
    if (seen[full]) return; seen[full] = 1;                         // identical in every field: counted once
    var cat = catMap[(it.fam || '') + '|' + (it.grp || '')] || fpGuessCat(it.fam, it.grp);
    var bcU = it.bc.toUpperCase();
    var i = idx[bcU + '|' + String(it.pack || '').toUpperCase() + '|' + it.name.toUpperCase()];
    if (i === undefined) {
      /* the same barcode and packing whose name the POS changed, or one that was added in the sheet */
      var cands = (byBc[bcU] || []).filter(function (j) {
        var r = rows[j];
        return !touched[j] && (String(r[D.PACK - 1] || '').toUpperCase() === String(it.pack || '').toUpperCase() || r[D.SRC - 1] === 'added in sheet');
      });
      if (cands.length) i = cands[0];
    }
    if (i === undefined || touched[i]) {
      rows.push([it.bc, fpKey(it.bc), it.name, it.brand, it.code, it.pack, it.unit, it.fam, it.grp, cat, 'pos', 'map', 'POS', today]);
      touched[rows.length - 1] = 1; added++; return;
    }
    touched[i] = 1;
    var r = rows[i], before = r.slice(0, D.SRC).join('\u0001');
    r[D.BC - 1] = it.bc; r[D.KEY - 1] = fpKey(it.bc);
    if (r[D.NAMEBY - 1] !== 'hand') r[D.NAME - 1] = it.name;       // a name corrected in the sheet is never overwritten
    r[D.BRAND - 1] = it.brand; r[D.CODE - 1] = it.code; r[D.PACK - 1] = it.pack; r[D.UNIT - 1] = it.unit;
    r[D.FAM - 1] = it.fam; r[D.GRP - 1] = it.grp;
    if (r[D.CATBY - 1] !== 'hand') r[D.CAT - 1] = cat;             // neither is a category chosen in the sheet
    var wasSheet = r[D.SRC - 1] === 'added in sheet';
    r[D.SRC - 1] = 'POS';
    if (r.slice(0, D.SRC).join('\u0001') !== before || wasSheet) { r[D.UPD - 1] = today; changed++; } else same++;
  });
  var kept = 0; rows.forEach(function (r, i) { if (!touched[i]) kept++; });
  return {rows: rows, added: added, changed: changed, same: same, kept: kept};
}

/* ======================= reading the database ======================= */
function fpDbRowCount(db) { return Math.max(0, db.getLastRow() - 1); }
function fpFindIn(db, col, text, n) {
  if (!text || !n) return [];
  var hits = db.getRange(2, col, n, 1).createTextFinder(text).matchEntireCell(true).findAll();
  return hits.map(function (h) { var rr = h.getRow(); return {row: rr, vals: db.getRange(rr, 1, 1, CFG.DBH.length).getValues()[0]}; });
}
/** the exact barcode first; only if nothing matches, the barcode without its leading zeros */
function fpFind(db, bcText) {
  var n = fpDbRowCount(db); if (!n || !bcText) return null;
  var exact = fpFindIn(db, CFG.D.BC, bcText, n);
  if (exact.length) return fpPick(exact);
  var k = fpKey(bcText);
  return k ? fpPick(fpFindIn(db, CFG.D.KEY, k, n)) : null;
}
/** for a big paste: read the database once */
function fpLoadIndex(db) {
  var n = fpDbRowCount(db), ix = {bc: {}, key: {}};
  if (!n) return ix;
  var v = db.getRange(2, 1, n, CFG.DBH.length).getValues();
  for (var i = 0; i < n; i++) {
    var h = {row: i + 2, vals: v[i]};
    var b = fpText(v[i][CFG.D.BC - 1]).toUpperCase(), k = fpText(v[i][CFG.D.KEY - 1]).toUpperCase();
    (ix.bc[b] = ix.bc[b] || []).push(h); if (k) (ix.key[k] = ix.key[k] || []).push(h);
  }
  return ix;
}
function fpFromIndex(ix, bcText) {
  var b = fpText(bcText).toUpperCase();
  if (ix.bc[b]) return fpPick(ix.bc[b]);
  var k = fpKey(bcText).toUpperCase();
  return ix.key[k] ? fpPick(ix.key[k]) : null;
}
function fpToday() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
/** the heading row, trimmed */
function fpHeads(sh) {
  return sh.getRange(1, 1, 1, CFG.HEAD.length).getValues()[0].map(function (x) { return String(x).trim(); });
}
/** an offer tab: SL, then PRODUCT NAME and BARCODE (in either order — older tabs had BARCODE first) */
function fpIsEntry(sh, heads) {
  if (!sh) return false;
  var nm = sh.getName();
  if (nm === CFG.DB || nm === CFG.MAP || nm === CFG.LOG || nm === CFG.GUIDE || nm === CFG.BR) return false;
  var h = heads || sh.getRange(1, 1, 1, 3).getValues()[0].map(function (x) { return String(x).trim(); });
  return h[0] === 'SL' && ((h[1] === 'PRODUCT NAME' && h[2] === 'BARCODE') || (h[1] === 'BARCODE' && h[2] === 'PRODUCT NAME'));
}
/** the tab has all the headings, but in an older order */
function fpNeedsLayout(heads) {
  var H = CFG.HEAD, cur = heads.slice(0, H.length);
  if (cur.join('|') === H.join('|')) return false;
  return H.every(function (x) { return cur.indexOf(x) >= 0; });
}
function fpColLetter(c) { var s = ''; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }
/** columns into the order of CFG.HEAD — with their data, dropdowns and colours. true if anything moved. */
function fpEnsureLayout(sh) {
  var heads = fpHeads(sh);
  if (!fpNeedsLayout(heads)) return false;
  var H = CFG.HEAD, cur = heads.slice(0, H.length);
  var frozen = sh.getFrozenColumns();
  if (frozen) sh.setFrozenColumns(0);                /* a column cannot be moved across the frozen edge */
  for (var t = 0; t < H.length; t++) {
    var p = cur.indexOf(H[t]);
    if (p === t) continue;                           /* p > t: everything left of t is already in place */
    var L = fpColLetter(p + 1);
    sh.moveColumns(sh.getRange(L + ':' + L), t + 1);
    cur.splice(p, 1); cur.splice(t, 0, H[t]);
  }
  if (frozen) sh.setFrozenColumns(frozen);
  return true;
}
/** the entry rows of a tab: from row 2 to its last row (rows may have been inserted or deleted) */
function fpRows(sh) { return Math.max(1, sh.getMaxRows() - CFG.FIRST + 1); }
function fpSlFormula(r) { return '=IF(COUNTA(B' + r + ':C)=0,"",ROW()-1)'; }
/** after rows were inserted or deleted: SL numbers, the barcode format and the CATEGORY list on every row */
function fpFixRows(sh) {
  var n = fpRows(sh), C = CFG.C, f = [];
  for (var i = 0; i < n; i++) f.push([fpSlFormula(CFG.FIRST + i)]);
  sh.getRange(CFG.FIRST, C.SL, n, 1).setFormulas(f);
  sh.getRange(CFG.FIRST, C.BC, n, 1).setNumberFormat('@');
  var cats = sh.getRange(CFG.FIRST, C.CAT, n, 1).getDataValidations(), dv = null;
  for (var j = 0; j < n && !dv; j++) dv = cats[j][0];
  if (dv) sh.getRange(CFG.FIRST, C.CAT, n, 1).setDataValidation(dv);
}
function fpNewDbRow(bc, name, cat, today) {
  return [bc, fpKey(bc), name, '', '', '', '', '', '', cat || '', 'hand', cat ? 'hand' : '', 'added in sheet', today];
}
function fpAppendDb(db, rows) {
  if (!rows.length) return;
  var start = db.getLastRow() + 1;
  db.getRange(start, 1, rows.length, 2).setNumberFormat('@');
  db.getRange(start, 1, rows.length, CFG.DBH.length).setValues(rows);
}

/* ======================= the barcode cache (for speed) ======================= */
/*
 * Opening the master and searching its DATABASE is the slow part of a scan. So
 * the DATABASE is also kept in Google's script cache, cut into buckets: a
 * barcode typed in a branch sheet is found without opening the master at all.
 * It is built again every 5 hours, after a PRODUCT MASTER update, after an edit
 * in the DATABASE tab, and when a bucket has dropped out of the cache.
 */
var FP_TTL = 21600;
function fpCacheVer() { return PropertiesService.getScriptProperties().getProperty('CACHE_V') || '0'; }
function fpCacheReset() {
  var p = PropertiesService.getScriptProperties();
  p.setProperty('CACHE_V', String(Number(p.getProperty('CACHE_V') || 0) + 1));
}
function fpHash(s) { var h = 5381; for (var i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0; return h; }
function fpCacheKeys(bc) {
  var b = fpText(bc).toUpperCase(), k = fpKey(bc).toUpperCase(), out = [];
  if (b) out.push('B' + b);
  if (k) out.push('K' + k);
  return out;
}
/** the whole DATABASE into the cache; returns the buckets as well */
function fpCacheBuild(db) {
  var ix = fpLoadIndex(db), ent = {}, cnt = 0, D = CFG.D;
  var put = function (key, hits) {
    var h = fpPick(hits); if (!h) return;
    ent[key] = [h.row, h.vals[D.NAME - 1], h.vals[D.BRAND - 1], h.vals[D.PACK - 1], h.vals[D.CAT - 1]]; cnt++;
  };
  Object.keys(ix.bc).forEach(function (b) { if (b) put('B' + b, ix.bc[b]); });
  Object.keys(ix.key).forEach(function (k) { if (k) put('K' + k, ix.key[k]); });
  var nb = Math.max(8, Math.ceil(cnt / 400)), buckets, ok = false;
  while (!ok) {
    buckets = {};
    for (var i = 0; i < nb; i++) buckets[i] = {};
    Object.keys(ent).forEach(function (key) { buckets[fpHash(key) % nb][key] = ent[key]; });
    ok = Object.keys(buckets).every(function (i) { return JSON.stringify(buckets[i]).length < 95000; });
    if (!ok) nb *= 2;
  }
  var fc = {v: fpCacheVer(), nb: nb, b: buckets};
  try { fpCacheSave(fc, Object.keys(buckets)); } catch (err) {}
  return fc;
}
function fpCacheSave(fc, ids) {
  var cache = CacheService.getScriptCache(), pre = 'fp' + fc.v + ':', out = {};
  ids.forEach(function (i) { out[pre + i] = JSON.stringify(fc.b[i]); });
  out[pre + 'm'] = String(fc.nb);
  var keys = Object.keys(out);
  for (var s = 0; s < keys.length; s += 100) {
    var part = {}; keys.slice(s, s + 100).forEach(function (k) { part[k] = out[k]; });
    cache.putAll(part, FP_TTL);
  }
}
/** the buckets these barcodes fall in, from the cache — null if the cache is not there (then build it) */
function fpCacheLoad(bcs) {
  try {
    var cache = CacheService.getScriptCache(), v = fpCacheVer(), pre = 'fp' + v + ':';
    var nb = Number(cache.get(pre + 'm') || 0);
    if (!nb) return null;
    var want = {};
    bcs.forEach(function (bc) { fpCacheKeys(bc).forEach(function (k) { want[fpHash(k) % nb] = 1; }); });
    var ids = Object.keys(want), got = cache.getAll(ids.map(function (i) { return pre + i; })), b = {};
    for (var j = 0; j < ids.length; j++) {
      var raw = got[pre + ids[j]];
      if (!raw) return null;                         /* a bucket dropped out: build again */
      b[ids[j]] = JSON.parse(raw);
    }
    return {v: v, nb: nb, b: b};
  } catch (err) { return null; }
}
/** a hit in the same shape as fpFind: {row, vals} */
function fpCacheFind(fc, bc) {
  var keys = fpCacheKeys(bc), D = CFG.D;
  for (var i = 0; i < keys.length; i++) {
    var bk = fc.b[fpHash(keys[i]) % fc.nb], e = bk && bk[keys[i]];
    if (!e) continue;
    var vals = []; for (var c = 0; c < CFG.DBH.length; c++) vals.push('');
    vals[D.BC - 1] = fpText(bc); vals[D.NAME - 1] = e[1]; vals[D.BRAND - 1] = e[2]; vals[D.PACK - 1] = e[3]; vals[D.CAT - 1] = e[4];
    return {row: e[0], vals: vals};
  }
  return null;
}
/** after a write into the DATABASE: the same change in the cache (called inside the lock) */
function fpCachePatch(changes) {             // changes: [{bc, row, name?, cat?, add?: [row, name, brand, pack, cat]}]
  if (!changes.length) return;
  var fc = fpCacheLoad(changes.map(function (c) { return c.bc; }));
  if (!fc) return;                             /* not in the cache: it is built again from the DATABASE anyway */
  var touched = {};
  changes.forEach(function (c) {
    fpCacheKeys(c.bc).forEach(function (key, idx) {
      var id = fpHash(key) % fc.nb, bk = fc.b[id], e = bk[key];
      if (c.add) { if (!e || idx === 0) { bk[key] = c.add.slice(); touched[id] = 1; } return; }
      if (e && e[0] === c.row) {
        if (c.name !== undefined) e[1] = c.name;
        if (c.cat !== undefined) e[4] = c.cat;
        touched[id] = 1;
      }
    });
  });
  var ids = Object.keys(touched);
  if (ids.length) try { fpCacheSave(fc, ids); } catch (err) {}
}
/** time trigger + menu: build the cache again */
function warmCache() {
  var db = fpMaster().getSheetByName(CFG.DB);
  if (!db) return;
  fpCacheReset();
  fpCacheBuild(db);
}

/* ======================= the edit trigger (runs as the owner) ======================= */
function handleEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() === CFG.DB) { fpCacheReset(); return; }   /* the DATABASE was typed in by hand: cache out of date */
  var C = CFG.C, D = CFG.D, H = CFG.HEAD;
  var total = fpRows(sh);
  var grid = sh.getRange(1, 1, total + 1, H.length).getValues();     /* the heading and the whole tab: one read */
  var heads = grid[0].map(function (x) { return String(x).trim(); });
  if (!fpIsEntry(sh, heads)) return;
  var r0 = e.range.getRow(), nr = e.range.getNumRows(), c0 = e.range.getColumn(), nc = e.range.getNumColumns();
  var cols = {};
  for (var c = c0; c < c0 + nc; c++) cols[c] = 1;
  if (fpNeedsLayout(heads)) {                          /* a tab in an older column order: put it in the new one first */
    var ml = LockService.getScriptLock();
    if (!ml.tryLock(25000)) return;
    try { fpEnsureLayout(sh); } finally { ml.releaseLock(); }
    var moved = {};
    Object.keys(cols).forEach(function (k) { var p = H.indexOf(heads[k - 1]); moved[p >= 0 ? p + 1 : k] = 1; });
    cols = moved;
    grid = sh.getRange(1, 1, total + 1, H.length).getValues();
  }
  var hit = function (c) { return !!cols[c]; };
  var first = Math.max(r0, CFG.FIRST), last = Math.min(r0 + nr - 1, CFG.FIRST + total - 1);
  if (last < first) return;
  if (hit(C.SL)) {                                     /* SL was typed over or cleared: its formula back */
    var slf = []; for (var q = first; q <= last; q++) slf.push([fpSlFormula(q)]);
    sh.getRange(first, C.SL, slf.length, 1).setFormulas(slf);
  }
  var bcHit = hit(C.BC), nameHit = hit(C.NAME), catHit = hit(C.CAT);
  var priceHit = hit(C.CP) || hit(C.SP) || hit(C.OFFER);
  if (!bcHit && !nameHit && !catHit && !priceHit) return;

  var all = grid.slice(1);
  var n = last - first + 1, off = first - CFG.FIRST;
  var today = fpToday();

  /* the master is opened only when it is needed */
  var mid = '';
  try { mid = PropertiesService.getScriptProperties().getProperty('MASTER_ID') || ''; } catch (err) {}
  if (!FP_MASTER && e.source && (!mid || e.source.getId() === mid)) FP_MASTER = e.source;
  var db = null;
  var dbGet = function () { if (!db) db = fpMaster().getSheetByName(CFG.DB); return db; };

  /* the barcodes of these rows: from the cache; only if it is not there, from the DATABASE (and the cache is built) */
  var fc = null;
  if (bcHit || nameHit || catHit) {
    var bcs = [];
    for (var b = 0; b < n; b++) { var t = fpText(all[off + b][C.BC - 1]); if (t) bcs.push(t); }
    if (bcs.length) {
      fc = fpCacheLoad(bcs);
      if (!fc) { if (!dbGet()) return; fc = fpCacheBuild(db); }
    }
  }
  var find = function (bc) { return fc ? fpCacheFind(fc, bc) : null; };

  var appendRows = [], dbSet = [];            // dbSet: {bc, row, col, value}
  for (var i = 0; i < n; i++) {
    var v = all[off + i];
    var bc = fpText(v[C.BC - 1]), name = String(v[C.NAME - 1] || '').trim(), cat = String(v[C.CAT - 1] || '').trim();
    var state = String(v[C.STATE - 1] || '');
    var outName = v[C.NAME - 1], outCat = v[C.CAT - 1], outBrand = v[C.BRAND - 1], outPack = v[C.PACK - 1], outState = state;
    var typedName = nameHit && name, typedCat = catHit && cat;
    if (bcHit) {
      if (!bc) {
        if (state === 'db' && !typedName) { outName = ''; outCat = ''; }   // only what the script had put there
        outBrand = ''; outPack = '';
        outState = (typedName || (name && state !== 'db')) ? 'nobc' : '';
      } else {
        var h = find(bc);
        if (h) {
          var dv = h.vals;
          if (typedName) { if (name.toUpperCase() !== String(dv[D.NAME - 1]).toUpperCase()) dbSet.push({bc: bc, row: h.row, col: D.NAME, value: name}); }
          else outName = dv[D.NAME - 1];
          if (typedCat) { if (cat !== String(dv[D.CAT - 1])) dbSet.push({bc: bc, row: h.row, col: D.CAT, value: cat}); }
          else outCat = dv[D.CAT - 1];
          outBrand = dv[D.BRAND - 1]; outPack = dv[D.PACK - 1]; outState = 'db';
        } else {
          if (state === 'db' && !typedName) { outName = ''; name = ''; }
          if (state === 'db' && !typedCat) { outCat = ''; cat = ''; }
          outBrand = ''; outPack = '';
          if (name) { appendRows.push(fpNewDbRow(bc, name, cat, today)); outState = 'added'; }
          else outState = 'new';
        }
      }
    } else if ((nameHit || catHit) && bc) {
      /* a hand change on a row that has a barcode: into the database */
      var h2 = find(bc);
      if (h2) {
        if (nameHit && name && name.toUpperCase() !== String(h2.vals[D.NAME - 1]).toUpperCase())
          dbSet.push({bc: bc, row: h2.row, col: D.NAME, value: name});
        if (catHit && cat && cat !== String(h2.vals[D.CAT - 1]))
          dbSet.push({bc: bc, row: h2.row, col: D.CAT, value: cat});
        outState = 'db';
      } else if (name) {
        var pending = appendRows.some(function (a) { return fpText(a[0]) === bc; });
        if (!pending) appendRows.push(fpNewDbRow(bc, name, cat, today));
        outState = 'added';
      } else outState = 'new';
    } else if (nameHit && !bc) {
      outState = name ? 'nobc' : '';
      if (!name) { outBrand = ''; outPack = ''; }
    }
    v[C.NAME - 1] = outName; v[C.CAT - 1] = outCat; v[C.BRAND - 1] = outBrand; v[C.PACK - 1] = outPack; v[C.STATE - 1] = outState;
  }

  /* CHECK from what is in memory now; written for the edited rows, and elsewhere only where it changed */
  var labels = fpCheckLabels(all), rest = [];
  for (var j = 0; j < total; j++) {
    if (j >= off && j < off + n) continue;
    if (labels[j] !== String(all[j][C.CHECK - 1] || '')) rest.push(j);
  }
  var block = all.slice(off, off + n);
  if (bcHit) sh.getRange(first, C.NAME, n, 1).setValues(block.map(function (r) { return [r[C.NAME - 1]]; }));
  /* CATEGORY, BRAND, PACKING, CHECK, ROW STATE sit side by side: one write */
  sh.getRange(first, C.CAT, n, C.STATE - C.CAT + 1).setValues(block.map(function (r, k) {
    return [r[C.CAT - 1], r[C.BRAND - 1], r[C.PACK - 1], labels[off + k], r[C.STATE - 1]];
  }));
  if (rest.length > 20) sh.getRange(CFG.FIRST, C.CHECK, total, 1).setValues(labels.map(function (l) { return [l]; }));
  else rest.forEach(function (k) { sh.getRange(CFG.FIRST + k, C.CHECK).setValue(labels[k]); });

  /* only writing into the database opens the master and waits for the other branches */
  if (!dbSet.length && !appendRows.length) return;
  if (!dbGet()) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;
  try {
    var changes = [];
    dbSet.forEach(function (w) {
      var row = w.row;
      if (fpText(db.getRange(row, D.BC).getValue()).toUpperCase() !== fpText(w.bc).toUpperCase()) {   /* the DATABASE moved since */
        var f = fpFind(db, w.bc); if (!f) return; row = f.row;
      }
      var by = w.col === D.NAME ? D.NAMEBY : D.CATBY;
      db.getRange(row, w.col).setValue(w.value);
      db.getRange(row, by).setValue('hand');
      db.getRange(row, D.UPD).setValue(today);
      changes.push(w.col === D.NAME ? {bc: w.bc, row: row, name: w.value} : {bc: w.bc, row: row, cat: w.value});
    });
    var fresh = appendRows.filter(function (a) { return !fpFind(db, fpText(a[0])); });   /* another branch may have just added it */
    var start = db.getLastRow() + 1;
    fpAppendDb(db, fresh);
    fresh.forEach(function (a, k) { changes.push({bc: a[0], add: [start + k, a[D.NAME - 1], '', '', a[D.CAT - 1]]}); });
    fpCachePatch(changes);
  } finally {
    lock.releaseLock();
  }
}
/* ================ the change trigger: a row inserted or deleted (runs as the owner) ================ */
function handleChange(e) {
  var t = e && e.changeType;
  if (t !== 'INSERT_ROW' && t !== 'REMOVE_ROW') return;
  var ss = (e && e.source) || SpreadsheetApp.getActive();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;
  try {
    ss.getSheets().forEach(function (sh) {
      if (!fpIsEntry(sh)) return;
      fpEnsureLayout(sh);
      fpFixRows(sh);
      if (t === 'INSERT_ROW') fpWarnRanges(sh);     /* rows added at the very bottom are covered as well */
      fpRecheck(sh);
    });
  } finally {
    lock.releaseLock();
  }
}
/** the CHECK label of every row: ⚠ for the same line twice, otherwise the row's state */
function fpCheckLabels(v) {
  var C = CFG.C;
  var dup = fpDuplicates(v.map(function (r) { return [r[C.BC - 1], r[C.NAME - 1], r[C.CAT - 1], r[C.CP - 1], r[C.SP - 1], r[C.OFFER - 1]]; }));
  return v.map(function (r, i) {
    if (!fpText(r[C.BC - 1]) && !String(r[C.NAME - 1] || '').trim()) return '';
    if (dup[i]) return '⚠ Same as SL ' + dup[i];
    return fpStateLabel(String(r[C.STATE - 1] || ''));
  });
}
/** the CHECK column for the whole tab */
function fpRecheck(sh) {
  var n = fpRows(sh), v = sh.getRange(CFG.FIRST, 1, n, CFG.HEAD.length).getValues();
  sh.getRange(CFG.FIRST, CFG.C.CHECK, n, 1).setValues(fpCheckLabels(v).map(function (l) { return [l]; }));
}

/* ======================= owner tools ======================= */
function fpIsOwner() {
  try {
    var o = SpreadsheetApp.getActive().getOwner();
    return !o || o.getEmail() === Session.getEffectiveUser().getEmail();
  } catch (err) { return true; }
}
function fpOwnerOnly(p) {
  var me = Session.getEffectiveUser();
  p.addEditor(me);
  var others = p.getEditors().filter(function (u) { return u.getEmail() !== me.getEmail(); });
  if (others.length) p.removeEditors(others);
  if (p.canDomainEdit()) p.setDomainEdit(false);
}
function fpProtectWhole(s) {
  s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  fpOwnerOnly(s.protect().setDescription('Offer sheet — owner only'));
}
/**
 * An offer tab. A whole-sheet lock stops everyone but the owner from inserting or
 * deleting rows — even with BARCODE to NOTE left open, a new row also touches the
 * locked columns. So only the heading row is locked (that also keeps the columns in
 * place); SL and BRAND → ROW STATE only warn, and the script fills them again.
 */
function fpProtectEntry(s) {
  fpEnsureLayout(s);
  s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  s.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (/^Offer sheet/.test(p.getDescription())) p.remove();
  });
  fpOwnerOnly(s.getRange(1, 1, 1, s.getMaxColumns()).protect().setDescription('Offer sheet — the heading row'));
  fpWarnRanges(s);
  fpFixRows(s);
  s.hideColumns(CFG.C.STATE);
}
/** SL and BRAND → ROW STATE: filled by the script. Anyone may still insert or delete a row across them. */
function fpWarnRanges(s) {
  s.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (p.getDescription() === 'Offer sheet — filled by the script') p.remove();
  });
  var C = CFG.C, n = fpRows(s);
  [s.getRange(CFG.FIRST, C.SL, n, 1), s.getRange(CFG.FIRST, C.BRAND, n, C.STATE - C.BRAND + 1)].forEach(function (r) {
    r.protect().setDescription('Offer sheet — filled by the script').setWarningOnly(true);
  });
}
function fpLog(ss, what, detail) {
  var lg = ss.getSheetByName(CFG.LOG); if (!lg) return;
  lg.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), what, detail]);
}

function setup() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (!fpIsOwner()) { ui.alert('Only the owner of this file can run the set-up.'); return; }
  if (!ss.getSheetByName(CFG.DB)) { ui.alert('This file has no DATABASE tab — run the set-up in the master sheet.'); return; }
  PropertiesService.getScriptProperties().setProperty('MASTER_ID', ss.getId());
  FP_MASTER = ss;
  fpEnsureTrigger(ss);
  if (!ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'warmCache'; }))
    ScriptApp.newTrigger('warmCache').timeBased().everyHours(5).create();   /* the cache lasts 6 hours at most */
  fpEnsureBranchTab(ss);
  var relinked = 0;
  fpBranchList().forEach(function (b) {
    try {
      if (!fpIsLinked(b.id)) { fpEnsureTrigger(b.id); relinked++; }
      SpreadsheetApp.openById(b.id).getSheets().forEach(function (s) { if (fpIsEntry(s)) fpProtectEntry(s); });   /* rows can be inserted there too */
    } catch (err) {}
  });
  [CFG.DB, CFG.MAP, CFG.LOG, CFG.BR].forEach(function (nm) { var s = ss.getSheetByName(nm); if (s) { fpProtectWhole(s); s.hideSheet(); } });
  var g = ss.getSheetByName(CFG.GUIDE); if (g) fpProtectWhole(g);
  var tabs = 0; ss.getSheets().forEach(function (s) { if (fpIsEntry(s)) { fpProtectEntry(s); tabs++; } });
  fpLog(ss, 'Set up', 'Edit trigger on; ' + tabs + ' offer tab(s) locked; back-office tabs hidden and locked' +
    (relinked ? ('; ' + relinked + ' branch sheet(s) re-linked') : '') + '.');
  ui.alert('Set-up done ✓',
    'This is the MASTER. Barcodes fill themselves here and in every branch sheet made from here.\n\n' +
    'Next: Offer Tools → Create a branch sheet — one for each branch. Do not use "Make a copy" for a branch.' +
    (relinked ? ('\n\n' + relinked + ' branch sheet(s) were re-linked.') : ''), ui.ButtonSet.OK);
}

/* ---------------------------------------------------------------- branch sheets */
function createBranchSheet() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (!fpIsOwner()) { ui.alert('Only the owner of this file can create branch sheets.'); return; }
  if (!ss.getSheetByName(CFG.DB)) { ui.alert('Branch sheets are made from the master (the file with the DATABASE).'); return; }
  PropertiesService.getScriptProperties().setProperty('MASTER_ID', ss.getId()); FP_MASTER = ss;
  var r = ui.prompt('Create a branch sheet', 'Name of the branch (for example: F5 — Al Wakrah):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var name = String(r.getResponseText() || '').trim(); if (!name) return;
  if (fpBranchList().some(function (b) { return b.name.toLowerCase() === name.toLowerCase(); })) {
    ui.alert('There is already a branch sheet called “' + name + '”. Offer Tools → Show the branch sheets.'); return;
  }
  var tmpl = ss.getSheets().filter(fpIsEntry)[0];
  if (!tmpl) { ui.alert('The master has no OFFER tab to build the branch sheet from.'); return; }
  var title = String(ss.getName()).replace(/\s*[—-]\s*MASTER\s*$/i, '') + ' — ' + name;
  var ns = SpreadsheetApp.create(title);
  var offer = tmpl.copyTo(ns).setName('OFFER');
  var guide = ss.getSheetByName(CFG.GUIDE); if (guide) guide.copyTo(ns).setName(CFG.GUIDE);
  ns.getSheets().forEach(function (s) { if (s.getName() !== 'OFFER' && s.getName() !== CFG.GUIDE) ns.deleteSheet(s); });
  offer.getRange(CFG.FIRST, CFG.C.BC, fpRows(offer), CFG.HEAD.length - 1).clearContent();   /* a clean start, whatever the master held */
  fpProtectEntry(offer);
  var g2 = ns.getSheetByName(CFG.GUIDE); if (g2) fpProtectWhole(g2);
  fpEnsureTrigger(ns);
  var shared = true;
  try { DriveApp.getFileById(ns.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT); } catch (err) { shared = false; }
  try { var par = DriveApp.getFileById(ss.getId()).getParents(); if (par.hasNext()) DriveApp.getFileById(ns.getId()).moveTo(par.next()); } catch (err2) {}
  var link = ns.getUrl() + '#gid=' + offer.getSheetId();
  fpEnsureBranchTab(ss).appendRow([name, link, ns.getId(), fpToday()]);
  fpLog(ss, 'Branch sheet created', name + ' · ' + link);
  fpShowLink(name, link, shared);
}
function fpShowLink(name, link, shared) {
  var safe = String(link).replace(/"/g, '&quot;');
  var html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial,sans-serif;line-height:1.5">' +
    '<p><b>' + name.replace(/</g, '&lt;') + '</b> has its own sheet now.</p>' +
    '<p><a href="' + safe + '" target="_blank">Open it</a> — or copy the link:</p>' +
    '<input style="width:100%;padding:6px;font:13px monospace" value="' + safe + '" onclick="this.select()" readonly>' +
    '<p>' + (shared ? 'It is shared as <b>Anyone with the link → Editor</b>: send this link to the branch, and paste it into the flyer app for that branch.'
                    : 'Share it yourself: open it → Share → Anyone with the link → Editor. Then send the link to the branch and paste it into the flyer app.') + '</p>' +
    '<p style="color:#666">Barcodes typed there fill themselves from this master. Nothing else to set up.</p></div>'
  ).setWidth(560).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Branch sheet ready ✓');
}
function showBranchSheets() {
  var list = fpBranchList(), ui = SpreadsheetApp.getUi();
  if (!list.length) { ui.alert('No branch sheets yet — Offer Tools → Create a branch sheet.'); return; }
  var html = '<div style="font:14px Arial,sans-serif;line-height:1.6">' + list.map(function (b) {
    var ok = fpIsLinked(b.id);
    return '<p><b>' + b.name.replace(/</g, '&lt;') + '</b> ' + (ok ? '✓' : '⚠ not linked — run 1. Set up') +
      '<br><a href="' + b.link.replace(/"/g, '&quot;') + '" target="_blank">' + b.link.replace(/</g, '&lt;') + '</a></p>';
  }).join('') + '</div>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(Math.min(520, 110 + list.length * 70)), 'Branch sheets');
}
/** after a database update: every branch sheet fills its scanned rows again */
function refillAllBranches() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (!fpIsOwner()) { ui.alert('Only the owner of this file can do this.'); return; }
  var ix = fpLoadIndex(fpMaster().getSheetByName(CFG.DB)), done = [], total = 0;
  var files = [{name: 'this master', ss: ss}].concat(fpBranchList().map(function (b) {
    try { return {name: b.name, ss: SpreadsheetApp.openById(b.id)}; } catch (err) { return {name: b.name, ss: null}; }
  }));
  files.forEach(function (f) {
    if (!f.ss) { done.push(f.name + ': could not be opened'); return; }
    var n = 0; f.ss.getSheets().forEach(function (sh) { if (fpIsEntry(sh)) n += fpRefillSheet(sh, ix); });
    total += n; done.push(f.name + ': ' + n + ' row(s)');
  });
  fpLog(ss, 'Branch sheets filled again', done.join(' · '));
  ui.alert('Done ✓', done.join('\n'), ui.ButtonSet.OK);
}

function fpFindMasterTab(ss) {
  var ours = [CFG.DB, CFG.MAP, CFG.LOG, CFG.GUIDE, CFG.BR];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    if (ours.indexOf(s.getName()) >= 0 || fpIsEntry(s)) continue;
    var nr = Math.min(80, s.getLastRow()), ncol = Math.min(40, s.getLastColumn());
    if (nr < 1 || ncol < 1) continue;
    var top = s.getRange(1, 1, nr, ncol).getValues();
    for (var r = 0; r < top.length; r++) {
      var low = top[r].map(function (v) { return String(v).trim().toLowerCase(); });
      if (low.indexOf('barcode') >= 0 && low.indexOf('description') >= 0) return s;
    }
  }
  return null;
}
function fpReadMap(map) {
  var m = {}; if (!map) return m;
  var n = Math.max(0, map.getLastRow() - 1); if (!n) return m;
  map.getRange(2, 1, n, 3).getValues().forEach(function (r) { if (r[2]) m[String(r[0]) + '|' + String(r[1])] = String(r[2]); });
  return m;
}

function updateFromMaster() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (!fpIsOwner()) { ui.alert('Only the owner of this file can update the database.'); return; }
  var src = fpFindMasterTab(ss);
  if (!src) {
    ui.alert('No PRODUCT MASTER tab found',
      'First bring it in: File → Import → Upload → PRODUCT_MASTER.xls → Import location: "Insert new sheet(s)" → Import data.\n\nThen run this again.',
      ui.ButtonSet.OK);
    return;
  }
  var ok = ui.alert('Update the database from the tab “' + src.getName() + '”?',
    'New products are added, changed ones are updated. Names and categories corrected in the sheet are kept. ' +
    'Nothing is deleted. The imported tab is removed afterwards.', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  var lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    var items = fpParseMaster(src.getDataRange().getValues());
    if (items.length < 10) { ui.alert('Only ' + items.length + ' products were found in that tab — nothing was changed.'); return; }
    var db = ss.getSheetByName(CFG.DB), map = ss.getSheetByName(CFG.MAP);
    var catMap = fpReadMap(map), newGroups = [], seenG = {};
    items.forEach(function (it) {
      var k = (it.fam || '') + '|' + (it.grp || '');
      if (!catMap[k] && !seenG[k]) { seenG[k] = 1; var g = fpGuessCat(it.fam, it.grp); catMap[k] = g; newGroups.push([it.fam, it.grp, g, '']); }
    });
    if (newGroups.length && map) map.getRange(map.getLastRow() + 1, 1, newGroups.length, 4).setValues(newGroups);
    var n = fpDbRowCount(db);
    var rows = n ? db.getRange(2, 1, n, CFG.DBH.length).getValues() : [];
    var res = fpMerge(rows, items, catMap, fpToday());
    var R = res.rows, CH = 4000;
    db.getRange(2, 1, R.length, 2).setNumberFormat('@');
    for (var i = 0; i < R.length; i += CH) db.getRange(2 + i, 1, Math.min(CH, R.length - i), CFG.DBH.length).setValues(R.slice(i, i + CH));
    var msg = 'New: ' + res.added + '   Changed: ' + res.changed + '   Unchanged: ' + res.same +
              '   Kept (not in this master): ' + res.kept + (newGroups.length ? ('   New groups in the map: ' + newGroups.length) : '');
    fpLog(ss, 'Database updated from “' + src.getName() + '”', items.length + ' lines read. ' + msg);
    ss.deleteSheet(src);
    fpCacheReset(); fpCacheBuild(db);
    ui.alert('Database updated ✓', msg + (newGroups.length ? '\n\nCheck the new groups at the bottom of CATEGORY MAP.' : '') +
      '\n\nEvery branch sheet uses the new database from now on. To refresh rows that were already scanned: Offer Tools → Fill every branch sheet again.', ui.ButtonSet.OK);
  } finally { lock.releaseLock(); }
}

function addBranchTab() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (!fpIsOwner()) { ui.alert('Only the owner of this file can add a tab.'); return; }
  var tmpl = ss.getSheets().filter(fpIsEntry)[0];
  if (!tmpl) { ui.alert('There is no offer tab to copy.'); return; }
  var r = ui.prompt('Add a branch tab', 'Name of the new tab (for example: F5 — Al Wakrah):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var name = String(r.getResponseText() || '').trim(); if (!name) return;
  if (ss.getSheetByName(name)) { ui.alert('A tab called “' + name + '” already exists.'); return; }
  var t = tmpl.copyTo(ss).setName(name);
  t.getRange(CFG.FIRST, CFG.C.BC, fpRows(t), CFG.HEAD.length - 1).clearContent();
  fpProtectEntry(t);
  t.activate();
  ui.alert('Tab “' + name + '” is ready ✓', 'Copy its link (with gid=) into the flyer app for that branch.', ui.ButtonSet.OK);
}

/** after a database update: fill again every row the script had filled (hand-typed rows stay as they are) */
function recheckActiveTab() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet(), ui = SpreadsheetApp.getUi();
  if (!fpIsEntry(sh)) { ui.alert('Open an offer tab first.'); return; }
  var refreshed = fpRefillSheet(sh, fpLoadIndex(fpMaster().getSheetByName(CFG.DB)));
  ui.alert('Checked ✓', refreshed + ' row(s) filled again from the database.', ui.ButtonSet.OK);
}
function fpRefillSheet(sh, ix) {
  fpEnsureLayout(sh);
  var C = CFG.C, D = CFG.D, n = fpRows(sh);
  var v = sh.getRange(CFG.FIRST, 1, n, CFG.HEAD.length).getValues(), refreshed = 0;
  var cn = [], cc = [], ij = [], st = [];
  v.forEach(function (r) {
    var bc = fpText(r[C.BC - 1]), state = String(r[C.STATE - 1] || '');
    var out = [r[C.NAME - 1], r[C.CAT - 1], r[C.BRAND - 1], r[C.PACK - 1], state];
    if (bc && (state === 'db' || state === 'new')) {
      var h = fpFromIndex(ix, bc);
      if (h) { out = [h.vals[D.NAME - 1], h.vals[D.CAT - 1], h.vals[D.BRAND - 1], h.vals[D.PACK - 1], 'db']; refreshed++; }
    }
    cn.push([out[0]]); cc.push([out[1]]); ij.push([out[2], out[3]]); st.push([out[4]]);
  });
  sh.getRange(CFG.FIRST, C.NAME, n, 1).setValues(cn);
  sh.getRange(CFG.FIRST, C.CAT, n, 1).setValues(cc);
  sh.getRange(CFG.FIRST, C.BRAND, n, 2).setValues(ij);
  sh.getRange(CFG.FIRST, C.STATE, n, 1).setValues(st);
  fpRecheck(sh);
  return refreshed;
}

function warmCacheMenu() {
  warmCache();
  SpreadsheetApp.getUi().alert('Done ✓', 'The barcode cache was built again from the DATABASE.', SpreadsheetApp.getUi().ButtonSet.OK);
}
function reapplyCategories() {
  var ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  if (!fpIsOwner()) { ui.alert('Only the owner of this file can do this.'); return; }
  var db = ss.getSheetByName(CFG.DB), map = fpReadMap(ss.getSheetByName(CFG.MAP)), D = CFG.D;
  var n = fpDbRowCount(db); if (!n) return;
  var v = db.getRange(2, D.CAT, n, 1).getValues(), fg = db.getRange(2, D.FAM, n, 2).getValues(), by = db.getRange(2, D.CATBY, n, 1).getValues();
  var changed = 0;
  for (var i = 0; i < n; i++) {
    if (by[i][0] === 'hand') continue;
    var c = map[String(fg[i][0]) + '|' + String(fg[i][1])];
    if (c && c !== v[i][0]) { v[i][0] = c; changed++; }
  }
  db.getRange(2, D.CAT, n, 1).setValues(v);
  fpCacheReset(); fpCacheBuild(db);
  fpLog(ss, 'Category map applied', changed + ' product(s) moved to another category.');
  ui.alert('Done ✓', changed + ' product(s) moved to another category. Categories chosen by hand in the sheet were left alone.', ui.ButtonSet.OK);
}
