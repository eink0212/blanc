/**
 * Blanc - White Wine Cellar / バックエンド
 *
 * フロントエンドは GitHub Pages（https://eink0212.github.io/blanc/）に移行済み。
 * この Code.gs は「ウェブアプリ」としてデプロイし、JSON API として使う。
 * 旧 HTMLサービス版の doGet / include は不要になったため削除してある。
 *
 * 追記した API レイヤーは既存コードの書き方に合わせ、バッククォートと
 * // 形式のコメントを使っていない。
 */

var SHEET_WINES    = 'セラー';
var SHEET_PURCHASE = '202603-購入履歴';   /* 旧。2026-10-04 以降の購入履歴は別ファイル（purchaseBook_）。fillWineIds 等の古い保守用関数だけが参照 */
var SHEET_TASTING  = 'テイスティング';

var HEADERS = {
  wines:    ['ID','名前','生産者','ヴィンテージ','色','容量(ml)','単価(¥)','在庫(本)','商品コード','仕入先','登録日','更新日','購入日'],
  purchase: ['ID','ワインID','ワイン名','生産者','ヴィンテージ','本数','単価(¥)','小計(¥)','仕入先','購入日'],
  tasting:  ['ID','ワインID','ワイン名','日付','シーン','評価(★)','外観-色調','外観-清澄','香り-印象','香り-詳細','味わい-甘辛','味わい-詳細','総合コメント','登録日']
};

var SHEET_DRINKING = '飲酒履歴';
var HEADERS_DRINKING = ['ID','ワインID','ワイン名','生産者','ヴィンテージ','飲んだ日','本数','単価(¥)','メモ','登録日'];

function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('SS_IDが設定されていません');
  return SpreadsheetApp.openById(id);
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1A1916').setFontColor('#C8A84B').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function uid() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}


function getAllData() {
  try {
    var ss = getSpreadsheet();
    return {
      wines:     sheetToObjects(getOrCreateSheet(ss, SHEET_WINES, HEADERS.wines), HEADERS.wines),
      purchases: purchasesForApp_(),
      tastings:  sheetToObjects(getOrCreateSheet(ss, SHEET_TASTING, HEADERS.tasting), HEADERS.tasting),
      drinking:  sheetToObjects(getOrCreateSheet(ss, SHEET_DRINKING, HEADERS_DRINKING), HEADERS_DRINKING)
    };
  } catch(e) {
    return { error: e.message };
  }
}


function getWinesOnly() {
  try {
    var ss = getSpreadsheet();
    return { wines: sheetToObjects(getOrCreateSheet(ss, SHEET_WINES, HEADERS.wines), HEADERS.wines) };
  } catch(e) { return { error: e.message }; }
}

function getSubData() {
  try {
    var ss = getSpreadsheet();
    return {
      purchases: purchasesForApp_(),
      tastings:  sheetToObjects(getOrCreateSheet(ss, SHEET_TASTING, HEADERS.tasting), HEADERS.tasting),
      drinking:  sheetToObjects(getOrCreateSheet(ss, SHEET_DRINKING, HEADERS_DRINKING), HEADERS_DRINKING)
    };
  } catch(e) { return { error: e.message }; }
}
/**
 * シートを圧縮形 { h: [列名...], r: [[値...], ...] } で返す。
 *
 * 従来は行ごとに列名を持つオブジェクトの配列を返していた。
 * 日本語の列名が全行ぶん繰り返されるため、購入履歴などでは
 * JSON の半分近くが列名の重複だった（実測 360KB）。
 * 列名を1回だけ送ることで転送量と解析時間を減らす。
 *
 * フロント側は新旧どちらの形式も受け取れるようにしてあるので、
 * 貼り替えの前後どちらの状態でも動く。
 */
function sheetToObjects(sheet, headers) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { h: headers, r: [] };
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = (row[0] === undefined || row[0] === null) ? '' : String(row[0]);
    if (!id || id === 'undefined' || id === 'null') continue;
    out.push(rowToStrings_(row, headers.length));
  }
  return { h: headers, r: out };
}

/** 1行を先頭 n 列ぶんの文字列配列にする（日付は yyyy-MM-dd） */
function rowToStrings_(row, n) {
  var rec = [];
  for (var j = 0; j < n; j++) {
    var v = row[j];
    if (v instanceof Date) {
      rec.push(Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd'));
    } else {
      rec.push((v === undefined || v === null) ? '' : String(v));
    }
  }
  return rec;
}

function saveWine(wine) {
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_WINES, HEADERS.wines);
    var now = new Date().toLocaleDateString('ja-JP');
    if (wine.id) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(wine.id)) {
          sheet.getRange(i+1, 1, 1, HEADERS.wines.length).setValues([[
            wine.id, wine.name||'', wine.producer||'', wine.vintage||'',
            wine.color||'白', wine.volume||750, wine.price||0, wine.stock||0,
            wine.code||'', wine.supplier||'',
            data[i][10], now, wine.purchaseDate||''
          ]]);
          return { ok: true };
        }
      }
    }
    var newId = uid();
    sheet.appendRow([
      newId, wine.name||'', wine.producer||'', wine.vintage||'',
      wine.color||'白', wine.volume||750, wine.price||0, wine.stock||0,
      wine.code||'', wine.supplier||'',
      now, now, wine.purchaseDate||''
    ]);
    return { ok: true, id: newId };
  } catch(e) { return { error: e.message }; }
}

function deleteWine(id) {
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_WINES, HEADERS.wines);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { ok: true }; }
    }
    return { error: '対象が見つかりません' };
  } catch(e) { return { error: e.message }; }
}

function adjustStock(params) {
  try {
    var id = params.id;
    var delta = params.delta;
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_WINES, HEADERS.wines);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var newStock = Math.max(0, (parseInt(data[i][7]) || 0) + delta);
        /* 従来は 14列目(N列)に書いていた。HEADERS.wines は13列までなので
           N列は誰も読まない幽霊列になり、getDataRange の範囲も1列広がっていた。
           更新日は L列(12)。H〜L をまとめて1回で書き、往復も1回に減らす。 */
        sheet.getRange(i+1, 8, 1, 5).setValues([[
          newStock, data[i][8], data[i][9], data[i][10],
          new Date().toLocaleDateString('ja-JP')
        ]]);
        return { ok: true, stock: newStock };
      }
    }
    return { error: '対象が見つかりません' };
  } catch(e) { return { error: e.message }; }
}

/* ==============================================================
 *  購入履歴（別ファイル「Ein's Wine 購入履歴」の年ごとのシート）
 *  セラー・飲酒履歴・テイスティングはこれまでどおりこのスプレッドシート。
 *  ファイルのIDはスクリプトプロパティ PURCHASE_SS_ID。
 * ============================================================== */

/** 年ごとのシートの列（2021年〜すべて共通。雛形シートも同じ） */
var BOOK_HEAD = ['購入日', '仕入先', '生産者', 'ワイン名', 'ヴィンテージ', '色', '本数', '単価(税込)', '小計', '購入者', '書類', 'メモ',
                 'ID', 'ワインID', '取込ID', '登録日時'];

/** 購入者（経費に入れるのは eins の分） */
var BUYERS = ['eins', '3rd'];

function purchaseBook_() {
  var id = PropertiesService.getScriptProperties().getProperty('PURCHASE_SS_ID');
  if (!id) throw new Error('PURCHASE_SS_IDが設定されていません');
  return SpreadsheetApp.openById(id);
}

/** 「2026年」などのシート。create なら、無いとき雛形をコピーして作る（年が変わったとき） */
function purchaseYearSheet_(book, year, create) {
  var name = year + '年', sh = book.getSheetByName(name);
  if (sh || !create) return sh;
  var tpl = book.getSheetByName('雛形');
  if (!tpl) throw new Error('購入履歴ファイルに雛形シートがありません');
  sh = tpl.copyTo(book).setName(name);
  sh.showSheet();
  book.setActiveSheet(sh);
  book.moveActiveSheet(2);                      /* 説明の次（新しい年が左） */
  return sh;
}

/** 年ごとのシートの全行（見出しを除く） */
function bookRows_(sh) {
  var n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, BOOK_HEAD.length).getValues() : [];
}

/** 年ごとのシートの1行を、アプリが読む旧形式（HEADERS.purchase の順）にする */
function toAppPurchase_(r) {
  return [r[12] || '', r[13] || '', r[3], r[2], r[4], r[6], r[7], r[8], r[1], r[0]];
}

/** アプリ（Historique の「買った」など）向けの購入履歴。去年と今年の分を旧形式の圧縮形で返す */
function purchasesForApp_() {
  var book = purchaseBook_(), y = new Date().getFullYear(), out = [];
  [y - 1, y].forEach(function (yr) {
    var sh = purchaseYearSheet_(book, yr, false);
    if (!sh) return;
    bookRows_(sh).forEach(function (r) {
      if (r[0] === '' || !String(r[3]).trim()) return;
      out.push(rowToStrings_(toAppPurchase_(r), HEADERS.purchase.length));
    });
  });
  return { h: HEADERS.purchase, r: out };
}

/** 購入履歴に行を足す。rows は {date:'yyyy-MM-dd', ...}。購入日の年のシートに振り分ける */
function appendPurchases_(book, rows) {
  var byYear = {};
  rows.forEach(function (p) {
    var m = String(p.date).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) throw new Error('購入日が不正です: ' + p.date);
    (byYear[m[1]] = byYear[m[1]] || []).push([
      new Date(+m[1], +m[2] - 1, +m[3]), p.supplier || '', p.producer || '', p.name || '', p.vintage || '', p.color || '白',
      p.qty, p.price, p.qty * p.price, p.buyer || '', '', p.memo || '', p.id, p.wineId || '', p.batchId || '', p.at || ''
    ]);
  });
  Object.keys(byYear).forEach(function (y) {
    var sh = purchaseYearSheet_(book, y, true);
    sh.getRange(sh.getLastRow() + 1, 1, byYear[y].length, BOOK_HEAD.length).setValues(byYear[y]);
  });
}

/** 旧API。アプリからは使っていないが、呼ばれた場合も購入履歴ファイルに書く */
function savePurchase(purchase) {
  try {
    var qty = parseInt(purchase.qty) || 1;
    var price = parseInt(purchase.price) || 0;
    appendPurchases_(purchaseBook_(), [{
      date: dateStr_(purchase.date) || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
      supplier: purchase.supplier, producer: purchase.producer, name: purchase.wineName, vintage: purchase.vintage,
      qty: qty, price: price, buyer: purchase.buyer, id: uid(), wineId: purchase.wineId,
      at: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
    }]);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

/** 照合用。前後の空白・連続空白・大文字小文字の違いを無視する */
function normKey_(v) {
  return String(v === undefined || v === null ? '' : v).trim().replace(/\s+/g, ' ').toLowerCase();
}

/** 購入履歴の同一判定キー（同じ日・同じ仕入先・同じワイン） */
function purchaseKey_(date, supplier, name, vintage) {
  return [date, supplier, name, vintage].map(normKey_).join('\u0001');
}

/** 購入日を yyyy-MM-dd の文字列にそろえる（シートの値が Date の場合がある） */
function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v === undefined || v === null ? '' : v).trim();
}

/**
 * 登録前の二重登録チェック（読み取りのみ）。
 * arg: { items: [{ date, supplier, name, vintage }] }
 * 戻り値: { matches: [ 各 item と同じ日・同じ仕入先・同じワインの既存購入履歴（無ければ空配列） ] }
 */
function checkReception(arg) {
  try {
    var items = (arg && arg.items) || [];
    var book = purchaseBook_(), index = {}, seen = {};
    items.forEach(function (it) {
      var y = String(it.date || '').slice(0, 4);
      if (!/^\d{4}$/.test(y) || seen[y]) return;
      seen[y] = true;
      var sh = purchaseYearSheet_(book, y, false);
      if (!sh) return;
      bookRows_(sh).forEach(function (r) {
        var k = purchaseKey_(dateStr_(r[0]), r[1], r[3], r[4]);
        (index[k] = index[k] || []).push({ id: String(r[12] || ''), qty: r[6], price: r[7] });
      });
    });
    return {
      ok: true,
      matches: items.map(function (it) {
        return index[purchaseKey_(it.date, it.supplier, it.name, it.vintage)] || [];
      })
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * 納品書（Reception）の一括取り込み。
 *
 * 従来はフロントから1件ごとに adjustStock / saveWine / savePurchase を順番に
 * 呼んでいたため、途中の1回が遅れると残りが登録されず、押し直すと先頭から
 * 二重登録になっていた。1回の呼び出しでまとめて書く。
 *
 * arg: { batchId: '...', items: [{ name, producer, vintage, supplier, date, qty, price, volume, code, buyer, color }] }
 *
 * - 購入履歴（別ファイルの「年」シート）を先に書き、そのあとセラーを更新する
 * - 同じ取り込みの中で「同じ日・同じ仕入先・同じワイン・同じ単価・同じ購入者」は本数をまとめて1行
 * - 単価・本数・購入日が無い行は登録しない（経費の合計が信用できなくなるため）
 * - 取込ID 列に batchId を書く。同じ取込ID が既にあれば何も書かずに
 *   { duplicate: true } を返す（通信が切れて押し直しても二重にならない）
 * - セラーは「名前 + ヴィンテージ + 生産者」が一致する行のうち一番下（最新）の行に
 *   在庫を足す。一致しなければ新しい行を作る
 * - 購入履歴の行は追加するだけで、既存行を消したり書き換えたりしない
 */
function importReception(arg) {
  var items = (arg && arg.items) || [];
  var batchId = String((arg && arg.batchId) || '');
  if (!batchId) return { error: '取込IDがありません' };
  if (!items.length) return { error: '取り込む行がありません' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { error: '他の保存処理が実行中です。少し待ってからもう一度押してください' };
  try {
    var ss = getSpreadsheet();
    var book = purchaseBook_();

    /* 同じ取込ID が既にあれば、前回の送信で登録済み（今回の購入日の年と今年のシートを見る） */
    var years = {};
    years[String(new Date().getFullYear())] = true;
    items.forEach(function (it) { var y = String(it.date || '').slice(0, 4); if (/^\d{4}$/.test(y)) years[y] = true; });
    for (var yk in years) {
      var shy = purchaseYearSheet_(book, yk, false);
      if (!shy || shy.getLastRow() < 2) continue;
      var ids = shy.getRange(2, 15, shy.getLastRow() - 1, 1).getValues();
      for (var q = 0; q < ids.length; q++) {
        if (String(ids[q][0]) === batchId) return { ok: true, duplicate: true };
      }
    }

    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var today = now.toLocaleDateString('ja-JP');

    /* 1. 検査と、同じ日・同じワイン・同じ単価・同じ購入者のまとめ */
    var results = [];
    var groups = [], gIndex = {};
    items.forEach(function (it, i) {
      var name = String(it.name || '').trim();
      var qty = parseInt(it.qty, 10) || 0;
      var price = parseInt(it.price, 10) || 0;
      var date = dateStr_(it.date);
      if (!name) { results[i] = { status: 'skip', reason: 'ワイン名がありません' }; return; }
      if (qty <= 0) { results[i] = { status: 'skip', reason: '本数がありません' }; return; }
      if (price <= 0) { results[i] = { status: 'skip', reason: '単価がありません' }; return; }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) { results[i] = { status: 'skip', reason: '購入日がありません' }; return; }
      var buyer = String(it.buyer || '').trim();
      if (buyer && BUYERS.indexOf(buyer) < 0) { results[i] = { status: 'skip', reason: '購入者が不明です: ' + buyer }; return; }
      var rec = {
        name: name, producer: String(it.producer || '').trim(), vintage: String(it.vintage || '').trim(),
        supplier: String(it.supplier || '').trim(), date: date, qty: qty, price: price,
        volume: parseInt(it.volume, 10) || 750, code: String(it.code || '').trim(), buyer: buyer,
        color: String(it.color || '').trim() || '白'
      };
      /* 購入者が違えば、同じ日・同じワインでも別の行にする（経費の区別のため） */
      var gk = purchaseKey_(rec.date, rec.supplier, rec.name, rec.vintage) + '\u0001' + normKey_(rec.producer) + '\u0001' + price + '\u0001' + buyer;
      if (gIndex[gk] === undefined) { gIndex[gk] = groups.length; groups.push({ rec: rec, members: [] }); }
      else groups[gIndex[gk]].rec.qty += qty;
      groups[gIndex[gk]].members.push(i);
      results[i] = { status: 'ok' };
    });
    if (!groups.length) {
      return { ok: true, purchasesAdded: 0, cellarNew: 0, cellarAdded: 0, cellarError: '', results: results,
               wines: { h: HEADERS.wines, r: [] }, purchases: { h: HEADERS.purchase, r: [] } };
    }

    /* 2. セラー側の行を決める（書き込みはまだ） */
    var wSheet = getOrCreateSheet(ss, SHEET_WINES, HEADERS.wines);
    var W = HEADERS.wines.length;
    var wData = wSheet.getDataRange().getValues();
    var wKey = function (name, vintage, producer) {
      return normKey_(name) + '\u0001' + normKey_(vintage) + '\u0001' + normKey_(producer);
    };
    var wIndex = {};
    for (var i = 1; i < wData.length; i++) {
      if (!wData[i][0]) continue;
      wIndex[wKey(wData[i][1], wData[i][3], wData[i][2])] = { row: wData[i].slice(0, W), sheetRow: i + 1, isNew: false };
    }
    var changed = [], appended = [];
    groups.forEach(function (g) {
      var r = g.rec;
      var k = wKey(r.name, r.vintage, r.producer);
      var hit = wIndex[k];
      if (hit) {
        hit.row[6] = r.price;
        hit.row[7] = Math.max(0, (parseInt(hit.row[7], 10) || 0) + r.qty);
        if (r.supplier) hit.row[9] = r.supplier;
        hit.row[11] = today;
        hit.row[12] = r.date;
        if (!hit.isNew && changed.indexOf(hit) < 0) changed.push(hit);
        g.cellar = 'add';
      } else {
        hit = { row: [uid(), r.name, r.producer, r.vintage, r.color, r.volume, r.price, r.qty,
                      r.code, r.supplier, today, today, r.date], isNew: true };
        wIndex[k] = hit;
        appended.push(hit);
        g.cellar = 'new';
      }
      g.wineId = hit.row[0];
      g.members.forEach(function (i) {
        results[i].cellar = g.cellar;
        results[i].merged = g.members.length > 1;
      });
    });

    /* 3. 購入履歴を先に書く（購入日の年のシートへ） */
    var pRecs = groups.map(function (g) {
      var r = g.rec;
      return { date: r.date, supplier: r.supplier, producer: r.producer, name: r.name, vintage: r.vintage, color: r.color,
               qty: r.qty, price: r.price, buyer: r.buyer, id: uid(), wineId: g.wineId, batchId: batchId, at: nowStr };
    });
    appendPurchases_(book, pRecs);
    SpreadsheetApp.flush();

    /* 4. セラー */
    var cellarError = '';
    try {
      changed.forEach(function (h) { wSheet.getRange(h.sheetRow, 1, 1, W).setValues([h.row]); });
      if (appended.length) {
        wSheet.getRange(wSheet.getLastRow() + 1, 1, appended.length, W)
          .setValues(appended.map(function (h) { return h.row; }));
      }
      SpreadsheetApp.flush();
    } catch (e) {
      cellarError = e.message;
    }

    var P = HEADERS.purchase.length;
    return {
      ok: true,
      purchasesAdded: pRecs.length,
      cellarNew: cellarError ? 0 : appended.length,
      cellarAdded: cellarError ? 0 : changed.length,
      cellarError: cellarError,
      results: results,
      wines: { h: HEADERS.wines, r: cellarError ? [] : changed.concat(appended).map(function (h) { return rowToStrings_(h.row, W); }) },
      purchases: { h: HEADERS.purchase, r: pRecs.map(function (p) {
        return rowToStrings_([p.id, p.wineId, p.name, p.producer, p.vintage, p.qty, p.price, p.qty * p.price, p.supplier, p.date], P);
      }) }
    };
  } catch (e) {
    return { error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveTasting(t) {
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_TASTING, HEADERS.tasting);
    var now = new Date().toLocaleDateString('ja-JP');
    var row = [null, t.wineId||'', t.wineName||'', t.date||'', t.scene||'', t.star||0, t.appearColor||'', t.appearClarity||'', t.noseFirst||'', t.noseDetail||'', t.palateSweet||'', t.palateDetail||'', t.comment||'', null];
    if (t.id) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(t.id)) {
          row[0] = t.id; row[13] = data[i][13];
          sheet.getRange(i+1, 1, 1, HEADERS.tasting.length).setValues([row]);
          return { ok: true };
        }
      }
    }
    var newId = uid(); row[0] = newId; row[13] = now;
    sheet.appendRow(row);
    return { ok: true, id: newId };
  } catch(e) { return { error: e.message }; }
}

function deleteTasting(id) {
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_TASTING, HEADERS.tasting);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { ok: true }; }
    }
    return { error: '対象が見つかりません' };
  } catch(e) { return { error: e.message }; }
}

function saveDrinking(d) {
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_DRINKING, HEADERS_DRINKING);
    var now = new Date().toLocaleDateString('ja-JP');
    sheet.appendRow([
      uid(), d.wineId||'', d.wineName||'', d.producer||'', d.vintage||'', d.date||now,
      d.qty||1, d.price||0, d.memo||'', now
    ]);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

function getNameMaster() {
  try {
    var ss = getSpreadsheet();
    var domaine = [];
    var wine = [];
    var ds = ss.getSheetByName('ドメーヌ変換マスター');
    if (ds) {
      var ddata = ds.getDataRange().getValues();
      for (var i = 1; i < ddata.length; i++) {
        var alias = String(ddata[i][0]||'').trim();
        var official = String(ddata[i][2]||'').trim();
        if (alias && official) domaine.push({alias: alias, official: official});
      }
    }
    var ws = ss.getSheetByName('ワイン名変換マスター');
    if (ws) {
      var wdata = ws.getDataRange().getValues();
      for (var j = 1; j < wdata.length; j++) {
        var walias = String(wdata[j][0]||'').trim();
        var wofficial = String(wdata[j][1]||'').trim();
        if (walias && wofficial) wine.push({alias: walias, official: wofficial});
      }
    }
    return {ok: true, domaine: domaine, wine: wine};
  } catch(e) { return {error: e.message}; }
}

function setup() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('SS_IDを設定してください');
  var ss = SpreadsheetApp.openById(id);
  getOrCreateSheet(ss, SHEET_WINES, HEADERS.wines);
  getOrCreateSheet(ss, SHEET_PURCHASE, HEADERS.purchase);
  getOrCreateSheet(ss, SHEET_TASTING, HEADERS.tasting);
  Logger.log('完了');
}

function callClaudeAPI(params) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if (!apiKey) return { error: 'ANTHROPIC_API_KEYが設定されていません' };
    var payload = {
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: params.mediaType, data: params.base64 } },
          { type: 'text', text: '\u3053\u306e\u7d0d\u54c1\u66f8\u753b\u50cf\u304b\u3089\u30ef\u30a4\u30f3\u306e\u8cfc\u5165\u30ea\u30b9\u30c8\u3092JSON\u914d\u5217\u3067\u629c\u5c11\u3002[{"name":"\u5546\u54c1\u540d","vintage":\u5e74\u6570\u304b\u306cnull,"volume":ml\u6570,"qty":\u672c\u6570,"price":\u7a0e\u8fbc\u5358\u4fa1,"code":"\u5546\u54c1\u30b3\u30fc\u30c9"}]\u30ef\u30a4\u30f3\u4ee5\u5916\u9664\u5916\u3002JSON\u306e\u307f\u8fd4\u3059\u3002' }
        ]
      }]
    };
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    var json = JSON.parse(response.getContentText());
    if (json.error) return { error: json.error.message };
    var text = (json.content || []).map(function(c) { return c.text || ''; }).join('');
    text = text.replace(/```json|```/g, '').trim();
    return { ok: true, result: text };
  } catch(e) {
    return { error: e.message };
  }
}

function fillMissingIds() {
  try {
    var ss = getSpreadsheet();
    var sheetNames = [SHEET_WINES, SHEET_PURCHASE, SHEET_TASTING, SHEET_DRINKING];
    var results = [];
    sheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (!sheet) { results.push(name + ': シートなし'); return; }
      var data = sheet.getDataRange().getValues();
      var col1 = [];
      var filled = 0;
      for (var i = 1; i < data.length; i++) {
        var idVal = String(data[i][0]).trim();
        if (!idVal || idVal === '' || idVal === 'undefined' || idVal === 'null') {
          col1.push([uid()]);
          filled++;
        } else {
          col1.push([data[i][0]]);
        }
      }
      if (filled > 0 && col1.length > 0) {
        sheet.getRange(2, 1, col1.length, 1).setValues(col1);
      }
      results.push(name + ': ' + filled + '件にID付与');
    });
    Browser.msgBox('ID付与完了\n\n' + results.join('\n'));
  } catch(e) {
    Browser.msgBox('エラー: ' + e.message);
  }
}

function fillWineIds() {
  try {
    var ss = getSpreadsheet();
    var wineData = ss.getSheetByName(SHEET_WINES).getDataRange().getValues();
    var nameToId = {};
    for (var i = 1; i < wineData.length; i++) {
      var wId = String(wineData[i][0]).trim();
      var wName = String(wineData[i][1]).trim();
      if (wId && wName) nameToId[wName] = wId;
    }
    var pSheet = ss.getSheetByName(SHEET_PURCHASE);
    var pData = pSheet.getDataRange().getValues();
    var pCol = [];
    var pFilled = 0;
    for (var j = 1; j < pData.length; j++) {
      var pId = String(pData[j][1]).trim();
      var pName = String(pData[j][2]).trim();
      if ((!pId || pId === '') && nameToId[pName]) {
        pCol.push([nameToId[pName]]); pFilled++;
      } else {
        pCol.push([pData[j][1]]);
      }
    }
    if (pCol.length > 0) pSheet.getRange(2, 2, pCol.length, 1).setValues(pCol);

    var tSheet = ss.getSheetByName(SHEET_TASTING);
    var tData = tSheet.getDataRange().getValues();
    var tCol = [];
    var tFilled = 0;
    for (var k = 1; k < tData.length; k++) {
      var tId = String(tData[k][1]).trim();
      var tName = String(tData[k][2]).trim();
      if ((!tId || tId === '') && nameToId[tName]) {
        tCol.push([nameToId[tName]]); tFilled++;
      } else {
        tCol.push([tData[k][1]]);
      }
    }
    if (tCol.length > 0) tSheet.getRange(2, 2, tCol.length, 1).setValues(tCol);

    var dSheet = ss.getSheetByName(SHEET_DRINKING);
    var dData = dSheet.getDataRange().getValues();
    var dCol = [];
    var dFilled = 0;
    for (var l = 1; l < dData.length; l++) {
      var dId = String(dData[l][1]).trim();
      var dName = String(dData[l][2]).trim();
      if ((!dId || dId === '') && nameToId[dName]) {
        dCol.push([nameToId[dName]]); dFilled++;
      } else {
        dCol.push([dData[l][1]]);
      }
    }
    if (dCol.length > 0) dSheet.getRange(2, 2, dCol.length, 1).setValues(dCol);

    Browser.msgBox('ワインID紐付け完了\n購入: ' + pFilled + '件\nテイスティング: ' + tFilled + '件\n飲酒: ' + dFilled + '件');
  } catch(e) {
    Browser.msgBox('エラー: ' + e.message);
  }
}

/* ==============================================================
 *  Web API レイヤー（GitHub Pages のフロントから呼ばれる）
 * ============================================================== */

/**
 * 合言葉。index.html の API_TOKEN と同じ文字列にすること。
 * 読み取りは合言葉なしで通し、書き込み系だけ必須にしている。
 *
 * この構成ではアクセス権を「自分のみ」にできない。
 * 「自分のみ」だと GAS は accounts.google.com へリダイレクトし、
 * そこには CORS ヘッダーが無いためブラウザが fetch を打ち切る。
 * よってデプロイは「全員」固定で、保護はこの合言葉で行う。
 */
var API_TOKEN = 'blanc-GSAIae2Dt40P';

/** doGet（JSONP フォールバック）で許可する読み取り専用の関数 */
var READ_ONLY_FNS = ['getAllData', 'getWinesOnly', 'getSubData', 'getNameMaster', 'checkReception'];

/** フロントから呼べる関数の一覧。ここに無い名前は実行されない。 */
function apiHandlers_() {
  return {
    getAllData:    function (arg) { return getAllData(); },
    getWinesOnly:  function (arg) { return getWinesOnly(); },
    getSubData:    function (arg) { return getSubData(); },
    getNameMaster: function (arg) { return getNameMaster(); },
    saveWine:      function (arg) { return saveWine(arg); },
    deleteWine:    function (arg) { return deleteWine(arg); },
    adjustStock:   function (arg) { return adjustStock(arg); },
    savePurchase:  function (arg) { return savePurchase(arg); },
    importReception: function (arg) { return importReception(arg); },
    checkReception:  function (arg) { return checkReception(arg); },
    saveTasting:   function (arg) { return saveTasting(arg); },
    deleteTasting: function (arg) { return deleteTasting(arg); },
    saveDrinking:  function (arg) { return saveDrinking(arg); }
  };
}

/**
 * POST 本体。
 * リクエスト: { fn: '関数名', arg: 引数, token: '（任意）' }
 * レスポンス: { ok: true, data: ... } / { ok: false, error: 'メッセージ' }
 *
 * ok:false は「呼び出しそのものが失敗した」場合だけ（旧 withFailureHandler 相当）。
 * 各関数が正常に返した { error: '...' } は ok:true の data に入る。
 * こうすることでフロント側の従来のエラー判定がそのまま動く。
 *
 * フロント側は Content-Type を text/plain で送ること。
 * application/json にすると CORS プリフライト（OPTIONS）が飛び、
 * GAS はこれに応答できないためリクエストが失敗する。
 */
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'リクエストの JSON を解析できませんでした' });
  }
  if (!body || typeof body !== 'object') {
    return jsonOut_({ ok: false, error: 'リクエストが不正です' });
  }
  return jsonOut_(dispatch_(body.fn, body.arg, body.token));
}

/**
 * GET。用途は 2 つ。
 *   1. 疎通確認 … デプロイ URL をブラウザで開くと { ok: true, data: 'pong' }
 *   2. JSONP フォールバック … ?fn=getAllData&callback=cb（読み取り関数のみ）
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var fn = p.fn || 'ping';
  var res;

  if (fn === 'ping') {
    res = { ok: true, data: 'pong' };
  } else if (READ_ONLY_FNS.indexOf(fn) === -1) {
    res = { ok: false, error: 'GET で実行できるのは読み取り関数のみです: ' + fn };
  } else {
    var arg;
    if (p.arg) {
      try { arg = JSON.parse(p.arg); } catch (err) { arg = p.arg; }
    }
    res = dispatch_(fn, arg, p.token);
  }

  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + JSON.stringify(res) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOut_(res);
}

function dispatch_(fn, arg, token) {
  var handlers = apiHandlers_();
  if (!fn || !handlers.hasOwnProperty(fn)) {
    return { ok: false, error: '未知の関数です: ' + fn };
  }
  /* 読み取りは素通し。書き込み系だけ合言葉を要求する */
  if (API_TOKEN && READ_ONLY_FNS.indexOf(fn) === -1 && token !== API_TOKEN) {
    return { ok: false, error: '認証に失敗しました' };
  }
  try {
    return { ok: true, data: handlers[fn](arg) };
  } catch (err) {
    return { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
