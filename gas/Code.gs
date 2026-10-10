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

/** 今日の日付（yyyy-MM-dd）。toLocaleDateString は "2026/10/8" 形式になり列の中で混ざっていた */
function today_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
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
    var now = today_();
    if (wine.id) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(wine.id)) {
          sheet.getRange(i+1, 1, 1, HEADERS.wines.length).setValues([[
            wine.id, wine.name||'', wine.producer||'', wine.vintage||'',
            wine.color||'白', wine.volume||750, wine.price||0, data[i][7],
            wine.code||'', wine.supplier||'',
            data[i][10], now, wine.purchaseDate||''
          ]]);
          /* 在庫は書き戻さない。編集画面を開いている間に別の端末や取り込みで
             本数が変わると、古い本数に戻っていた。在庫の増減は adjustStock だけで行う */
          return { ok: true, stock: parseInt(data[i][7], 10) || 0 };
        }
      }
      return { error: '対象のワインが見つかりません（削除された可能性があります）' };
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
          today_()
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
      p.qty, p.price, p.qty * p.price, p.buyer || '', p.doc || '', p.memo || '', p.id, p.wineId || '', p.batchId || '', p.at || ''
    ]);
  });
  Object.keys(byYear).forEach(function (y) {
    var sh = purchaseYearSheet_(book, y, true);
    sh.getRange(sh.getLastRow() + 1, 1, byYear[y].length, BOOK_HEAD.length).setValues(byYear[y]);
  });
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
    var today = today_();

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
        color: String(it.color || '').trim() || '白',
        toCellar: it.cellar !== false            /* cellar:false なら購入履歴だけ（赤ワインなどセラーで管理しないもの） */
      };
      /* 購入者が違えば、同じ日・同じワインでも別の行にする（経費の区別のため） */
      var gk = purchaseKey_(rec.date, rec.supplier, rec.name, rec.vintage) + '\u0001' + normKey_(rec.producer) + '\u0001' + price + '\u0001' + buyer + '\u0001' + rec.toCellar;
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
      if (!r.toCellar) {
        g.cellar = 'none'; g.wineId = '';
        g.members.forEach(function (i) { results[i].cellar = 'none'; results[i].merged = g.members.length > 1; });
        return;
      }
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
               qty: r.qty, price: r.price, buyer: r.buyer, id: uid(), wineId: g.wineId, batchId: batchId, at: nowStr,
               doc: String((arg && arg.doc) || '') };
    });
    try {
      appendPurchases_(book, pRecs);
      SpreadsheetApp.flush();
    } catch (e) {
      removeBatch_(book, years, batchId);
      return { error: '購入履歴に書き込めませんでした（何も登録していません）: ' + e.message };
    }

    /* 4. セラー。失敗したら購入履歴も取り消す（そのままだと押し直しが「登録済み」で弾かれ、セラーが直らなかった） */
    var cellarError = '';
    try {
      changed.forEach(function (h) { wSheet.getRange(h.sheetRow, 1, 1, W).setValues([h.row]); });
      if (appended.length) {
        wSheet.getRange(wSheet.getLastRow() + 1, 1, appended.length, W)
          .setValues(appended.map(function (h) { return h.row; }));
      }
      SpreadsheetApp.flush();
    } catch (e) {
      removeBatch_(book, years, batchId);
      return { error: 'セラーに書き込めませんでした。購入履歴も取り消したので、もう一度押してください: ' + e.message };
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
    var now = today_();
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
      return { error: '対象の記録が見つかりません（削除された可能性があります）' };
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
    var now = today_();
    var newId = uid();
    sheet.appendRow([
      newId, d.wineId||'', d.wineName||'', d.producer||'', d.vintage||'', d.date||now,
      d.qty||1, d.price||0, d.memo||'', now
    ]);
    return { ok: true, id: newId };
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

/* ==============================================================
 *  Web API レイヤー（GitHub Pages のフロントから呼ばれる）
 * ============================================================== */

/**
 * 取込ID の行の「書類」列に、保管した納品書のファイル名を入れる。
 * arg: { batchId, year, doc }
 */
/** 取込ID の行を購入履歴から消す（取り込みが途中で失敗したときの取り消し用） */
function removeBatch_(book, years, batchId) {
  for (var y in years) {
    var sh = purchaseYearSheet_(book, y, false);
    if (!sh || sh.getLastRow() < 2) continue;
    var ids = sh.getRange(2, 15, sh.getLastRow() - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === batchId) sh.deleteRow(i + 2);
  }
}

function setPurchaseDoc(arg) {
  var batchId = String((arg && arg.batchId) || ''), doc = String((arg && arg.doc) || '');
  if (!batchId || !doc) return { error: '取込IDとファイル名が必要です' };
  var sh = purchaseYearSheet_(purchaseBook_(), String(arg.year || new Date().getFullYear()), false);
  if (!sh || sh.getLastRow() < 2) return { error: 'その年のシートがありません' };
  var ids = sh.getRange(2, 15, sh.getLastRow() - 1, 1).getValues(), n = 0;
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === batchId) { sh.getRange(i + 2, 11).setValue(doc); n++; }
  return { ok: true, updated: n };
}

/**
 * 購入履歴の年ごとの本数・金額（トップ画面用）。購入者に関係なく全部数える。
 */
function getPurchaseSummary() {
  var years = [];
  purchaseBook_().getSheets().forEach(function (sh) {
    var m = sh.getName().match(/^(\d{4})年$/);
    if (!m) return;
    var b = 0, a = 0;
    bookRows_(sh).forEach(function (r) {
      if (!r[0] || !r[3]) return;
      if (typeof r[6] === 'number') b += r[6];
      if (typeof r[8] === 'number') a += r[8];
    });
    years.push({ year: +m[1], bottles: b, amount: a });
  });
  years.sort(function (p, q) { return p.year - q.year; });
  return { years: years };
}

/* ===== ワイン会（別ファイル「Ein's Wine ワイン会収支」。ファイルIDはスクリプトプロパティ EVENTS_SS_ID） =====
 * 開催 … 1回1行。収入・費用・利益は保存のたびにここで計算して書く（シートで見ても分かるように）
 * 参加者・費用 … 開催IDで紐づく行。保存のたびにその回の行を入れ替える
 * 集計 … 月ごとの合計（数式）
 */
var EV_HEAD = ['ID', '日付', '会の名前', '状態', '人数', '収入', '費用', '利益', '利益率', 'メモ', '料理', '登録日時', '更新日時'];
var EV_PEOPLE_HEAD = ['開催ID', '日付', '会の名前', '名前', '参加費'];
var EV_COST_HEAD = ['開催ID', '日付', '会の名前', '区分', '内容', '金額', 'セラーID', '生産者', 'ワイン名', 'ヴィンテージ'];
var EV_STATUS = ['実施', '予定', '中止'];
var EV_CATS = ['ワイン', '食費', '会場', 'その他'];

/** ワイン会のファイル。create が true なら無いとき作る（書き込みのときだけ） */
function eventsBook_(create) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('EVENTS_SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  if (!create) return null;
  var ss = SpreadsheetApp.create("Ein's Wine ワイン会収支");
  var DARK = '#1A1916', GOLD = '#C8A84B';
  var mk = function (sh, name, head, widths) {
    sh.setName(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setBackground(DARK).setFontColor(GOLD).setFontWeight('bold');
    sh.setFrozenRows(1);
    widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  };
  mk(ss.getSheets()[0], '開催', EV_HEAD, [130, 90, 160, 60, 50, 90, 90, 90, 70, 260, 260, 140, 140]);
  mk(ss.insertSheet(), '参加者', EV_PEOPLE_HEAD, [130, 90, 160, 140, 90]);
  mk(ss.insertSheet(), '費用', EV_COST_HEAD, [130, 90, 160, 70, 320, 90, 130, 180, 260, 70]);
  ss.getSheetByName('開催').getRange('B:B').setNumberFormat('yyyy/mm/dd');
  ss.getSheetByName('開催').getRange('F:H').setNumberFormat('¥#,##0');
  ss.getSheetByName('開催').getRange('I:I').setNumberFormat('0%');
  ss.getSheetByName('参加者').getRange('E:E').setNumberFormat('¥#,##0');
  ss.getSheetByName('費用').getRange('F:F').setNumberFormat('¥#,##0');
  var sum = ss.insertSheet('集計');
  sum.getRange(1, 1, 1, 6).setValues([['月', '回数', '収入', '費用', '利益', '利益率']]).setBackground(DARK).setFontColor(GOLD).setFontWeight('bold');
  sum.getRange('A2').setFormula('=SORT(UNIQUE(FILTER(TEXT(\'開催\'!B2:B,"yyyy-mm"),\'開催\'!B2:B<>"",\'開催\'!D2:D="実施")),1,FALSE)');
  for (var r = 2; r <= 60; r++) {
    sum.getRange(r, 2, 1, 5).setFormulas([[
      '=IF(A' + r + '="","",SUMPRODUCT((TEXT(\'開催\'!B$2:B,"yyyy-mm")=A' + r + ')*(\'開催\'!D$2:D="実施")))',
      '=IF(A' + r + '="","",SUMPRODUCT((TEXT(\'開催\'!B$2:B,"yyyy-mm")=A' + r + ')*(\'開催\'!D$2:D="実施"),\'開催\'!F$2:F))',
      '=IF(A' + r + '="","",SUMPRODUCT((TEXT(\'開催\'!B$2:B,"yyyy-mm")=A' + r + ')*(\'開催\'!D$2:D="実施"),\'開催\'!G$2:G))',
      '=IF(A' + r + '="","",C' + r + '-D' + r + ')',
      '=IF(OR(A' + r + '="",C' + r + '=0),"",E' + r + '/C' + r + ')'
    ]]);
  }
  sum.getRange('C:E').setNumberFormat('¥#,##0');
  sum.getRange('F:F').setNumberFormat('0%');
  sum.setFrozenRows(1);
  props.setProperty('EVENTS_SS_ID', ss.getId());
  return ss;
}

function evRows_(sh) {
  var n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, sh.getLastColumn()).getValues() : [];
}

/** ワイン会の一覧（参加者・費用つき）。件数は多くないので全部返す */
function getEvents() {
  var ss = eventsBook_(false);
  if (!ss) return { events: [] };
  var people = {}, costs = {};
  evRows_(ss.getSheetByName('参加者')).forEach(function (r) {
    if (!r[0]) return;
    (people[r[0]] = people[r[0]] || []).push({ name: String(r[3]), fee: Number(r[4]) || 0 });
  });
  evRows_(ss.getSheetByName('費用')).forEach(function (r) {
    if (!r[0]) return;
    (costs[r[0]] = costs[r[0]] || []).push({ cat: String(r[3]), desc: String(r[4]), amount: Number(r[5]) || 0,
      wineId: String(r[6] || ''), producer: String(r[7] || ''), name: String(r[8] || ''), vintage: String(r[9] || '') });
  });
  var events = evRows_(ss.getSheetByName('開催')).filter(function (r) { return r[0]; }).map(function (r) {
    return { id: String(r[0]), date: dateStr_(r[1]), name: String(r[2]), status: String(r[3] || '実施'),
      income: Number(r[5]) || 0, cost: Number(r[6]) || 0, profit: Number(r[7]) || 0,
      memo: String(r[9] || ''), dishes: String(r[10] || ''),
      people: people[r[0]] || [], costs: costs[r[0]] || [] };
  });
  events.sort(function (a, b) { return b.date < a.date ? -1 : b.date > a.date ? 1 : 0; });
  return { events: events };
}

/** シートから、その開催IDの行を消す（下から） */
function evDeleteRows_(sh, id) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return;
  var ids = sh.getRange(2, 1, n, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) sh.deleteRow(i + 2);
}

/**
 * 1回分を保存する。ev: { id?, date, name, status, memo, dishes, people:[{name,fee}], costs:[{cat,desc,amount,wineId,producer,name,vintage}] }
 * id が無ければ新規。参加者・費用はその回の行を丸ごと入れ替える。
 */
function saveEvent(ev) {
  if (!ev || !/^\d{4}-\d{2}-\d{2}$/.test(String(ev.date || ''))) return { error: '日付がありません' };
  if (!String(ev.name || '').trim()) return { error: '会の名前がありません' };
  var status = EV_STATUS.indexOf(ev.status) >= 0 ? ev.status : '実施';
  var people = (ev.people || []).filter(function (p) { return String(p.name || '').trim() || Number(p.fee); });
  var costs = (ev.costs || []).filter(function (c) { return String(c.desc || '').trim() || Number(c.amount); });
  var income = people.reduce(function (s, p) { return s + (Number(p.fee) || 0); }, 0);
  var cost = costs.reduce(function (s, c) { return s + (Number(c.amount) || 0); }, 0);
  var ss = eventsBook_(true), sh = ss.getSheetByName('開催');
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var m = ev.date.match(/^(\d{4})-(\d{2})-(\d{2})$/), d = new Date(+m[1], +m[2] - 1, +m[3]);
  var id = String(ev.id || ''), row = -1;
  if (id) {
    var v = evRows_(sh);
    for (var i = 0; i < v.length; i++) if (String(v[i][0]) === id) { row = i + 2; break; }
    if (row < 0) return { error: 'この回が見つかりません（削除された可能性があります）' };
  } else id = uid();
  var name = String(ev.name).trim();
  var rec = [id, d, name, status, people.length, income, cost, income - cost, income ? (income - cost) / income : '',
             String(ev.memo || ''), String(ev.dishes || ''), row > 0 ? sh.getRange(row, 12).getValue() : now, now];
  if (row > 0) sh.getRange(row, 1, 1, EV_HEAD.length).setValues([rec]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, EV_HEAD.length).setValues([rec]);
  var ps = ss.getSheetByName('参加者'), cs = ss.getSheetByName('費用');
  evDeleteRows_(ps, id); evDeleteRows_(cs, id);
  if (people.length) ps.getRange(ps.getLastRow() + 1, 1, people.length, EV_PEOPLE_HEAD.length).setValues(people.map(function (p) {
    return [id, d, name, String(p.name || '').trim(), Number(p.fee) || 0];
  }));
  if (costs.length) cs.getRange(cs.getLastRow() + 1, 1, costs.length, EV_COST_HEAD.length).setValues(costs.map(function (c) {
    return [id, d, name, EV_CATS.indexOf(c.cat) >= 0 ? c.cat : 'その他', String(c.desc || '').trim(), Number(c.amount) || 0,
            String(c.wineId || ''), String(c.producer || ''), String(c.name || ''), String(c.vintage || '')];
  }));
  return { ok: true, id: id, income: income, cost: cost, profit: income - cost };
}

function deleteEvent(id) {
  var ss = eventsBook_(false);
  if (!ss || !id) return { error: '対象が見つかりません' };
  var before = ss.getSheetByName('開催').getLastRow();
  ['開催', '参加者', '費用'].forEach(function (n) { evDeleteRows_(ss.getSheetByName(n), String(id)); });
  return before === ss.getSheetByName('開催').getLastRow() ? { error: '対象が見つかりません' } : { ok: true };
}

/* ===== 築地の仕入れ（ワイン会収支のファイルの「築地仕入れ」「店名」シート） =====
 * 1回の買い物の1店＝1行。どの会の分かは「会ID」で持つ（入力時に選ぶ。空なら未割り当て）
 * 区分: A＝アインズワイン / T＝てる（てるシェフに渡す分。会の食費は ¥10,000 で別に記録）/ 他
 */
var TJ_HEAD = ['ID', '日付', '店', '金額', '区分', '会ID', '会の名前', 'メモ', '登録日時'];

function tsukijiSheets_(ss) {
  var DARK = '#1A1916', GOLD = '#C8A84B';
  var sh = ss.getSheetByName('築地仕入れ');
  if (!sh) {
    sh = ss.insertSheet('築地仕入れ');
    sh.getRange(1, 1, 1, TJ_HEAD.length).setValues([TJ_HEAD]).setBackground(DARK).setFontColor(GOLD).setFontWeight('bold');
    sh.setFrozenRows(1);
    [130, 90, 150, 80, 50, 130, 160, 220, 140].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange('B:B').setNumberFormat('yyyy/mm/dd');
    sh.getRange('D:D').setNumberFormat('¥#,##0');
  }
  var ms = ss.getSheetByName('店名');
  if (!ms) {
    ms = ss.insertSheet('店名');
    ms.getRange(1, 1, 1, 3).setValues([['店名', '回数', '最後に行った日']]).setBackground(DARK).setFontColor(GOLD).setFontWeight('bold');
    ms.setFrozenRows(1);
    ms.setColumnWidth(1, 180); ms.setColumnWidth(3, 120);
  }
  return { rows: sh, shops: ms };
}

/** 店名シートの式（回数・最後に行った日）を付け直す */
function tjShopFormulas_(ms) {
  var n = ms.getLastRow() - 1;
  if (n < 1) return;
  var f = [];
  for (var r = 2; r <= n + 1; r++) {
    f.push(['=IF(A' + r + '="","",COUNTIF(\'築地仕入れ\'!C:C,A' + r + '))',
            '=IF(A' + r + '="","",IFERROR(MAXIFS(\'築地仕入れ\'!B:B,\'築地仕入れ\'!C:C,A' + r + '),""))']);
  }
  ms.getRange(2, 2, n, 2).setFormulas(f);
  ms.getRange(2, 3, n, 1).setNumberFormat('yyyy/mm/dd');
}

/** 築地の仕入れ全部と店の一覧（よく行く順） */
function getTsukiji() {
  var ss = eventsBook_(false);
  if (!ss || !ss.getSheetByName('築地仕入れ')) return { rows: [], shops: [] };
  var t = tsukijiSheets_(ss);
  var rows = evRows_(t.rows).filter(function (r) { return r[0]; }).map(function (r) {
    return { id: String(r[0]), date: dateStr_(r[1]), shop: String(r[2]), amount: Number(r[3]) || 0, tag: String(r[4] || ''),
             eventId: String(r[5] || ''), memo: String(r[7] || '') };
  });
  var cnt = {};
  rows.forEach(function (r) { if (r.date >= '2025-01-01') cnt[r.shop] = (cnt[r.shop] || 0) + 1; });
  var shops = evRows_(t.shops).map(function (r) { return String(r[0]); }).filter(function (s) { return s; });
  shops.sort(function (a, b) { return (cnt[b] || 0) - (cnt[a] || 0); });
  return { rows: rows, shops: shops };
}

/**
 * 仕入れを保存する。arg: { rows:[{id?, date, shop, amount, tag, eventId, memo}] }
 * id があればその行を書き換え、無ければ足す。知らない店は店名シートに足す。
 */
function saveTsukiji(arg) {
  var list = (arg && arg.rows) || [];
  if (!list.length) return { error: '保存する行がありません' };
  var ss = eventsBook_(true), t = tsukijiSheets_(ss);
  var names = {};
  evRows_(ss.getSheetByName('開催')).forEach(function (r) { names[String(r[0])] = String(r[2]); });
  var v = evRows_(t.rows), at = {};
  v.forEach(function (r, i) { at[String(r[0])] = i + 2; });
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  var add = [], ids = [], shopsKnown = {};
  evRows_(t.shops).forEach(function (r) { shopsKnown[String(r[0])] = true; });
  var newShops = [];
  for (var k = 0; k < list.length; k++) {
    var x = list[k], m = String(x.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    var shop = String(x.shop || '').trim();
    if (!m || !shop) return { error: '日付と店名が必要です（' + (k + 1) + '行目）' };
    var rec = [x.id || uid(), new Date(+m[1], +m[2] - 1, +m[3]), shop, Number(x.amount) || 0, String(x.tag || ''),
               String(x.eventId || ''), names[String(x.eventId || '')] || '', String(x.memo || ''), now];
    if (x.id) {
      if (!at[x.id]) return { error: '対象の行が見つかりません: ' + x.id };
      rec[8] = t.rows.getRange(at[x.id], 9).getValue() || now;
      t.rows.getRange(at[x.id], 1, 1, TJ_HEAD.length).setValues([rec]);
    } else add.push(rec);
    ids.push(rec[0]);
    if (!shopsKnown[shop]) { shopsKnown[shop] = true; newShops.push([shop]); }
  }
  if (add.length) t.rows.getRange(t.rows.getLastRow() + 1, 1, add.length, TJ_HEAD.length).setValues(add);
  if (newShops.length) {
    t.shops.getRange(t.shops.getLastRow() + 1, 1, newShops.length, 1).setValues(newShops);
    tjShopFormulas_(t.shops);
  }
  return { ok: true, ids: ids, newShops: newShops.length };
}

function deleteTsukiji(id) {
  var ss = eventsBook_(false);
  if (!ss || !ss.getSheetByName('築地仕入れ') || !id) return { error: '対象が見つかりません' };
  var sh = ss.getSheetByName('築地仕入れ'), before = sh.getLastRow();
  evDeleteRows_(sh, String(id));
  return before === sh.getLastRow() ? { error: '対象が見つかりません' } : { ok: true };
}

/* ===== 寺田倉庫（「寺田倉庫」シート。保管タブがどの端末でも同じ一覧を出すため） ===== */
var STORAGE_SHEET = '寺田倉庫';
var STORAGE_HEAD = ['寺田ID', '生産者', 'ワイン名', 'ヴィンテージ', '色', '産地', '容量', '入庫日', 'セラー在庫', 'セラーID', 'メモ'];

/** 見出し行（A列が「寺田ID」の行）の行番号。無ければ -1 */
function storageHeadRow_(sh) {
  var n = Math.min(10, sh.getLastRow());
  var a = n > 0 ? sh.getRange(1, 1, n, 1).getValues() : [];
  for (var i = 0; i < a.length; i++) if (String(a[i][0]).trim() === '寺田ID') return i + 1;
  return -1;
}

function getStorage() {
  var sh = getSpreadsheet().getSheetByName(STORAGE_SHEET);
  if (!sh) return { rows: [] };
  var hr = storageHeadRow_(sh);
  if (hr < 0 || sh.getLastRow() <= hr) return { rows: [] };
  var v = sh.getRange(hr + 1, 1, sh.getLastRow() - hr, STORAGE_HEAD.length).getDisplayValues();
  var rows = [];
  v.forEach(function (r) {
    if (!String(r[0]).trim()) return;
    rows.push({ wineId: r[0], producer: r[1], name: r[2], vintage: r[3], color: r[4], area: r[5],
                volume: parseInt(r[6], 10) || 750, inboundDate: r[7], cellarStock: r[8], cellarId: r[9], note: r[10] });
  });
  return { rows: rows, title: String(sh.getRange(1, 1).getValue()) };
}

/**
 * 寺田の WineList（アプリで読み込んだもの）でシートを作り直す。
 * メモ・セラー在庫・セラーID は同じ寺田IDの行から引き継ぐ。
 */
function saveStorage(arg) {
  var items = (arg && arg.rows) || [];
  if (!items.length) return { error: '寺田のリストが空です' };
  var sh = getSpreadsheet().getSheetByName(STORAGE_SHEET);
  if (!sh) return { error: '「' + STORAGE_SHEET + '」シートがありません' };
  var hr = storageHeadRow_(sh);
  if (hr < 0) return { error: '「' + STORAGE_SHEET + '」シートの見出し行が見つかりません' };
  var old = {};
  if (sh.getLastRow() > hr) {
    sh.getRange(hr + 1, 1, sh.getLastRow() - hr, STORAGE_HEAD.length).getValues().forEach(function (r) {
      if (r[0]) old[String(r[0])] = [r[8], r[9], r[10]];
    });
  }
  var TYPE = { 'White wine': '白', 'Red wine': '赤', 'Sweet wine': '甘口', 'Sparkling wine': '泡', 'Rose wine': 'ロゼ' };
  var out = items.map(function (w) {
    var id = String(w.wineId || '').trim(), keep = old[id] || ['', '', ''];
    return [id, String(w.producer || ''), String(w.name || ''), String(w.vintage || ''), TYPE[w.type] || String(w.type || ''),
            String(w.area || ''), parseInt(w.volume, 10) || 750, String(w.inboundDate || ''), keep[0], keep[1], keep[2]];
  });
  /* 先に消してから書くと、書き込みに失敗したときメモ等が消える。上書きしてから余った行だけ消す */
  out = out.filter(function (r) { return r[0]; });
  if (!out.length) return { error: '寺田IDのある行がありません' };
  var oldN = sh.getLastRow() - hr;
  sh.getRange(hr + 1, 1, out.length, 1).setNumberFormat('@');
  sh.getRange(hr + 1, 4, out.length, 1).setNumberFormat('@');
  sh.getRange(hr + 1, 1, out.length, STORAGE_HEAD.length).setValues(out);
  if (oldN > out.length) sh.getRange(hr + 1 + out.length, 1, oldN - out.length, STORAGE_HEAD.length).clearContent();
  sh.getRange(1, 1).setValue('寺田倉庫の在庫（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd') + ' の WineList から）');
  return { ok: true, count: out.length, storage: getStorage() };
}

/**
 * パスワード。スクリプトプロパティ APP_PASSWORD に保存する（コードは GitHub で公開されているので書かない）。
 * 読み取りも書き込みも、すべての呼び出しで必須。アプリは端末ごとに一度入力して保存する。
 *
 * この構成ではアクセス権を「自分のみ」にできない。
 * 「自分のみ」だと GAS は accounts.google.com へリダイレクトし、
 * そこには CORS ヘッダーが無いためブラウザが fetch を打ち切る。
 * よってデプロイは「全員」固定で、保護はこのパスワードで行う。
 */
function appPassword_() {
  return PropertiesService.getScriptProperties().getProperty('APP_PASSWORD') || '';
}

/** 読み取りだけの関数（書き込みのロックを取らない） */
var READ_ONLY_FNS = ['getAllData', 'getWinesOnly', 'getSubData', 'getNameMaster', 'checkReception', 'getStorage', 'getPurchaseSummary', 'checkPass', 'getEvents', 'getTsukiji'];

/** フロントから呼べる関数の一覧。ここに無い名前は実行されない。 */
function apiHandlers_() {
  return {
    checkPass:     function (arg) { return { ok: true }; },
    getAllData:    function (arg) { return getAllData(); },
    getWinesOnly:  function (arg) { return getWinesOnly(); },
    getSubData:    function (arg) { return getSubData(); },
    getNameMaster: function (arg) { return getNameMaster(); },
    saveWine:      function (arg) { return saveWine(arg); },
    deleteWine:    function (arg) { return deleteWine(arg); },
    adjustStock:   function (arg) { return adjustStock(arg); },
    importReception: function (arg) { return importReception(arg); },
    checkReception:  function (arg) { return checkReception(arg); },
    saveTasting:   function (arg) { return saveTasting(arg); },
    deleteTasting: function (arg) { return deleteTasting(arg); },
    saveDrinking:  function (arg) { return saveDrinking(arg); },
    getStorage:    function (arg) { return getStorage(); },
    getPurchaseSummary: function (arg) { return getPurchaseSummary(); },
    getEvents:     function (arg) { return getEvents(); },
    saveEvent:     function (arg) { return saveEvent(arg); },
    deleteEvent:   function (arg) { return deleteEvent(arg); },
    getTsukiji:    function (arg) { return getTsukiji(); },
    saveTsukiji:   function (arg) { return saveTsukiji(arg); },
    deleteTsukiji: function (arg) { return deleteTsukiji(arg); },
    saveStorage:   function (arg) { return saveStorage(arg); },
    setPurchaseDoc: function (arg) { return setPurchaseDoc(arg); }
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
 * それ以外は GET では受けない（パスワードが URL に残るため）。
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var fn = p.fn || 'ping';
  var res;

  if (fn === 'ping') {
    res = { ok: true, data: 'pong' };
  } else {
    res = { ok: false, error: 'GET では実行できません: ' + fn };
  }
  return jsonOut_(res);
}

function dispatch_(fn, arg, token) {
  var handlers = apiHandlers_();
  if (!fn || !handlers.hasOwnProperty(fn)) {
    return { ok: false, error: '未知の関数です: ' + fn };
  }
  /* すべての呼び出しでパスワードを確かめる。
     4桁なので総当たりされないよう、間違いが続いたら10分止める（全員共通） */
  var pass = appPassword_();
  if (!pass) return { ok: false, error: 'パスワードが設定されていません' };
  var cache = CacheService.getScriptCache();
  var fails = parseInt(cache.get('authFail') || '0', 10);
  if (fails >= 30) return { ok: false, error: 'パスワードの間違いが続いたため、10分ほど使えません' };
  if (String(token || '') !== pass) {
    cache.put('authFail', String(fails + 1), 600);
    return { ok: false, error: '認証に失敗しました' };
  }
  /* 書き込みは同時に走らせない。2台から同時に「飲んだ」を押すと、
     両方が同じ在庫を読んで片方の増減が消えていた。importReception は自前でロックする */
  var lock = null;
  if (READ_ONLY_FNS.indexOf(fn) === -1 && fn !== 'importReception') {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return { ok: false, error: '他の保存処理が実行中です。少し待ってからもう一度押してください' };
  }
  try {
    return { ok: true, data: handlers[fn](arg) };
  } catch (err) {
    return { ok: false, error: (err && err.message) ? err.message : String(err) };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
