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
var SHEET_PURCHASE = '202603-購入履歴';
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
      purchases: sheetToObjects(getOrCreateSheet(ss, SHEET_PURCHASE, HEADERS.purchase), HEADERS.purchase),
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
      purchases: sheetToObjects(getOrCreateSheet(ss, SHEET_PURCHASE, HEADERS.purchase), HEADERS.purchase),
      tastings:  sheetToObjects(getOrCreateSheet(ss, SHEET_TASTING, HEADERS.tasting), HEADERS.tasting),
      drinking:  sheetToObjects(getOrCreateSheet(ss, SHEET_DRINKING, HEADERS_DRINKING), HEADERS_DRINKING)
    };
  } catch(e) { return { error: e.message }; }
}
function sheetToObjects(sheet, headers) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      if (v instanceof Date) {
        obj[h] = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else {
        obj[h] = v !== undefined ? String(v) : '';
      }
    });
    return obj;
  }).filter(function(r) {
    return r['ID'] && r['ID'] !== '' && r['ID'] !== 'undefined' && r['ID'] !== 'null';
  });
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

function savePurchase(purchase) {
  try {
    var ss = getSpreadsheet();
    var sheet = getOrCreateSheet(ss, SHEET_PURCHASE, HEADERS.purchase);
    var qty = parseInt(purchase.qty) || 1;
    var price = parseInt(purchase.price) || 0;
    sheet.appendRow([
      uid(), purchase.wineId||'', purchase.wineName||'',
      purchase.producer||'', purchase.vintage||'',
      qty, price, qty*price,
      purchase.supplier||'',
      purchase.date||new Date().toLocaleDateString('ja-JP')
    ]);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
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
var READ_ONLY_FNS = ['getAllData', 'getWinesOnly', 'getSubData', 'getNameMaster'];

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
