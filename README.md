# Blanc - White Wine Cellar

白ワインセラーの在庫・購入・テイスティング管理アプリ。

- **フロントエンド** … GitHub Pages（このリポジトリの静的ファイル）
- **バックエンド** … Google Apps Script の Web API（`doPost`）
- **データ** … Google スプレッドシート `1J4kQGKk0SOjKcVbgmlDIspDRhVcOLZ9J_u-tN4MAWlY`（ファイル名「Blanc セラー管理」。セラー・棚卸し・飲酒履歴・テイスティング・名前マスター）
- **購入履歴** … 別ファイル「Ein's Wine 購入履歴」`1dine2pDyAc-GbHf0vHBy8qvQIv3_TByLxUOQDvYThJM`（2026-10-04〜）。
  年ごとのシート（「2026年」など。列は全年共通）＋仕入先・販売・集計・照合リスト。
  ID はスクリプトプロパティ `PURCHASE_SS_ID`。購入日の年のシートに書き、年が変われば「雛形」をコピーして「2027年」などを自動で作る。
  アプリ（getSubData）には去年と今年の分を旧形式（ID・ワインID・ワイン名…購入日）で返す。旧シート `202603-購入履歴` には書かない

公開URL: https://eink0212.github.io/blanc/

---

## ファイル構成

```
blanc/
├── index.html                 アプリ本体（HTML + CSS + JS 一体）
├── manifest.json              PWA 設定
├── sw.js                      Service Worker（オフライン・更新管理）
├── icon-192.png               アイコン
├── icon-512.png
├── icon-maskable-512.png      Android マスカブル用
├── apple-touch-icon.png       iOS ホーム画面用
├── .nojekyll                  GitHub Pages の Jekyll 処理を無効化
└── gas/
    └── Code.gs                GAS 側に貼り付けるバックエンド全文
```

`gas/Code.gs` は GitHub Pages では配信されない。GAS エディタへ手で貼るための控え。

---

## セットアップ手順

### 1. GAS 側

初回の移行時の手順。**現在の更新は clasp で行う**（「更新のしかた」参照）。

1. スプレッドシートの拡張機能 → Apps Script でエディタを開く。
2. 既存の `Code.gs` の中身を**全選択して消し**、`gas/Code.gs` の内容を
   まるごと貼り付ける。旧版との違いは次の 3 点だけ。
   - 旧HTMLサービス用の `doGet` と `include` を**削除**。フロントは GitHub Pages に
     移ったので不要。残したままだと新しい `doGet` と名前が衝突し、
     あとから定義されたほうに上書きされて API が動かなくなる。
   - 末尾に Web API レイヤー（`doPost` / `doGet` / `dispatch_` / `apiHandlers_`）を追加。
   - `getAllData` `saveWine` などの既存関数は**一切変更なし**。
3. 左メニューの `Index.html` は削除してよい（本体は GitHub Pages 側の `index.html`）。
4. 右上「デプロイ」→「デプロイを管理」→ 鉛筆アイコン →
   バージョン「**新バージョン**」→ デプロイ。
   - 次のユーザーとして実行: **自分**
   - アクセスできるユーザー: **全員**（「自分のみ」は動かない。セキュリティ節を参照）
   - 既存デプロイを更新すれば URL は変わらない。
     「新しいデプロイ」を作ると URL が変わるので、その場合は `index.html` の
     `GAS_URL` も書き換えること。
5. `/exec` URL をブラウザでそのまま開き、`{"ok":true,"data":"pong"}` が
   返ることを確認する。

### 2. フロントエンド側

`index.html` 冒頭の定数が、4. のデプロイ URL と一致しているか確認する。

```javascript
var GAS_URL = 'https://script.google.com/macros/s/＜デプロイID＞/exec';
```

移行前のデプロイ URL を初期値として入れてある。既存デプロイを更新した場合は
そのままでよい。

### 3. GitHub Pages

リポジトリの Settings → Pages → Source: `Deploy from a branch` /
Branch: `main` / フォルダ: `/ (root)`。数分で公開される。

### 4. スマホのホーム画面に追加

- **iOS (Safari)** … 共有 → 「ホーム画面に追加」
- **Android (Chrome)** … メニュー → 「アプリをインストール」

`display: standalone` なのでアドレスバーは表示されない。

---

## GAS 呼び出しの仕組み

GAS ウェブアプリ版で使っていた `google.script.run` は GitHub Pages では動かないので、
`gasCall` の中身だけを `fetch` に差し替えてある。**呼び出し側のコードは一行も
変えていない** — `gasCall('getAllData')` / `gasCall('saveWine', wine)` のままで、
戻り値の形も従来どおり。

| | 旧（google.script.run） | 新（fetch） |
|---|---|---|
| 関数が正常に返した | resolve（`{error:'...'}` もそのまま渡る） | 同じ |
| 関数が例外を投げた | reject | 同じ |
| 通信自体が失敗 | reject | 同じ |

サーバー側は `{ ok: true, data: ... }` / `{ ok: false, error: ... }` で包んで返し、
`gasCall` が開梱する。`ok:false` は「呼び出しそのものが失敗した」場合のみで、
各関数が正常に返した `{error:'...'}` は `ok:true` の `data` に入るため、
既存のエラー判定（`if (data.error) ...`）がそのまま機能する。

### レスポンスの形式

シートのデータは圧縮形で返る。列名を1回だけ送り、値は配列で並べる。

```json
{ "ok": true, "data": { "wines": {
    "h": ["ID", "名前", "生産者", "…"],
    "r": [["8abab9…", "Les Carmes de Rieussec", "Château Rieussec", "…"], …]
} } }
```

従来は行ごとに日本語の列名を持つオブジェクトの配列だったため、
JSON の約半分が列名の重複だった。実測値:

| | 旧形式 | 圧縮形 |
|---|---|---|
| `getWinesOnly` | 40,753 B | **21,049 B** |
| `getSubData` | 360,006 B | **179,708 B** |

フロント側の `toObjs()` が展開する。**旧形式もそのまま受け取れる**ので、
`Code.gs` の貼り替え前後どちらの状態でも動く。

なお GAS 自体の応答は1回あたり 2.5〜6秒かかる。起動を速くしているのは
主に `localStorage` のキャッシュで、2回目以降は通信を待たずに描画する。

---

### CORS について

GAS の `doPost` は CORS プリフライト（`OPTIONS`）に応答できない。
`Content-Type: application/json` で送るとプリフライトが発生して**必ず失敗する**。
`text/plain;charset=utf-8` で送ると単純リクエスト扱いになり、プリフライトなしで通る。
ボディの中身は JSON 文字列のままでよい。

```javascript
headers: { 'Content-Type': 'text/plain;charset=utf-8' }   // ← ここを変えないこと
```

万一 CORS で詰まった場合の保険として、`Code.gs` の `doGet` が
読み取り専用関数の JSONP 呼び出しに対応している
（`?fn=getAllData&callback=cb`）。

---

## 呼び出せる関数

| 関数名 | 引数 | 内容 | フロントで使用 |
|---|---|---|---|
| `getAllData` | なし | 全データ取得（wines/purchases/tastings/drinking） | －（分割ロードに移行） |
| `getWinesOnly` | なし | セラーデータのみ（高速起動用） | ○（起動時） |
| `getSubData` | なし | purchases/tastings/drinking を取得 | ○（起動後に裏で） |
| `getNameMaster` | なし | 変換マスター取得 | ○ |
| `saveWine` | wine オブジェクト | ワイン保存（編集では在庫を書かない。在庫は adjustStock だけ） | ○ |
| `deleteWine` | id 文字列 | ワイン削除 | ○ |
| `adjustStock` | `{id, delta}` | 在庫増減（飲んだ・寺田へ・±1本） | ○ |
| `importReception` | `{batchId, items:[...]}` | 納品書の一括取り込み（購入履歴 → セラー） | ○ |
| `checkReception` | `{items:[...]}` | 登録前の二重登録チェック（読み取りのみ） | ○ |
| `saveTasting` | tasting オブジェクト | テイスティング保存 | ○ |
| `deleteTasting` | id 文字列 | テイスティング削除 | ○ |
| `saveDrinking` | drinking オブジェクト | 飲酒履歴保存 | ○ |
| `getStorage` / `saveStorage` | なし / `{rows}` | 寺田倉庫シートの読み書き | ○ |
| `getPurchaseSummary` | なし | 購入履歴の年ごとの本数・金額（トップ画面） | ○ |
| `getEvents` / `saveEvent` / `deleteEvent` | なし / 1回分 / id | ワイン会（別ファイル。スクリプトプロパティ `EVENTS_SS_ID`。無ければ最初の保存で作る） | ○ |
| `getTsukiji` / `saveTsukiji` / `deleteTsukiji` | なし / `{rows}` / id | 築地の仕入れ（ワイン会収支の「築地仕入れ」「店名」シート） | ○ |
| `setPurchaseDoc` | `{batchId, year, doc}` | 取込IDの行の「書類」列を埋める | －（保守用） |

書き込み系は `dispatch_` で `LockService` を取り、1つずつ実行する（importReception は自前でロック）。

`Code.gs` の `apiHandlers_()` に載っていない関数は実行されない。
関数を増やしたら `apiHandlers_()` とこの表の両方に追加すること。

### Reception（納品書JSONの取り込み）

「登録する」は `importReception` を**1回だけ**呼ぶ。以前は1件ごとに
`adjustStock` / `saveWine` / `savePurchase` を順番に呼んでいたため、途中の1回が
遅れると残りが登録されず、押し直すと先頭から二重登録になっていた。

- **購入履歴を先に書き、そのあとセラー。** どちらかが失敗したら、その取込IDの購入履歴を
  取り消してエラーを返す（押し直せば最初からやり直せる）
- 同じ取り込みの中で「同じ日・同じ仕入先・同じワイン・同じ単価」は本数をまとめて1行
- 単価・本数・購入日の無い行は登録しない（スキップとして返す）
- 購入履歴シートの K・L 列に**取込ID・登録日時**を書く（見出しが無ければ自動で付ける）。
  同じ取込ID が既にあれば書き込まずに `{duplicate:true}` を返すので、通信が途切れて
  押し直しても二重にならない。アプリが読むのは従来どおり A〜J の10列
- セラーは**名前 + ヴィンテージ + 生産者**が一致する行のうち一番下（最新）の行に在庫を足す。
  一致しなければ新しい行を作る
- 登録前に `checkReception` で「同じ日・同じ仕入先・同じワイン」が既に購入履歴にあるか調べ、
  該当行は警告してチェックを外す。送信直前にももう一度調べ、該当があれば確認を出す
- `LockService` で同時書き込みを防ぐ
- **購入者**（eins / 3rd）を行ごとに選ぶ（一括変更あり。初期値は eins、JSON の `buyer` があればそれ）。
  購入履歴の「購入者」列に書く。列は見出しで探し、無ければ右端の次に作る（K〜M列の既存の見出しは触らない）。
  購入者が違えば同じ日・同じワインでも別の行。経費に入れるのは eins の分
- 購入履歴の行はアプリからは追加するだけで、消したり書き換えたりしない
  （「飲んだ」「削除」「寺田へ」を押しても購入履歴は変わらない）

---

## 更新のしかた

- **フロントを直した** … `index.html` を編集して push。加えて `sw.js` の
  `CACHE_VERSION` を上げる（`blanc-v1` → `blanc-v2`）。上げ忘れると
  古いキャッシュが残って更新が反映されない。
  画面の一番下の版表示（`index.html` の `appVersion`）も同じ番号・日付に揃える。
- **GAS を直した** … clasp で反映する（`.clasp.json` がこのフォルダにある。対象は `gas/` の2ファイルだけ）。
  ```
  clasp push                                   # gas/Code.gs と appsscript.json をアップロード
  clasp create-version "変更内容"               # 版を作る（表示された版番号を控える）
  clasp redeploy AKfycbw3Ay_CQ524GsfQ4NZWp5t5A0vYKvMLjI3io4y69ZQNpLc4Ow3njycvVymDZxSfESBj -V <版番号>
  ```
  **既存デプロイを更新する**（新規デプロイを作るとURLが変わる）。
  `clasp push` は GAS 側にあってここに無いファイルを消すので、GAS エディタで直接ファイルを足さないこと。
  反映の順番は GAS → フロント（push）。

---

## セキュリティ

### アクセス権は「全員」しか選べない

GAS のアクセス権を「自分のみ」にすると、GAS は accounts.google.com へ
リダイレクトを返す。そこには CORS ヘッダーが無いため、ブラウザは
クロスオリジンの fetch を打ち切る。**この構成で「自分のみ」は原理的に成立しない。**
旧 HtmlService 版にあった Google ログインという認証層は、移行によって失われている。

### パスワードによる保護

すべての呼び出し（読み取りも書き込みも）にパスワードが要る。

- パスワードは GAS のスクリプトプロパティ `APP_PASSWORD` にだけ置く（リポジトリは公開なのでコードに書かない）
- アプリは端末ごとに一度入力させ、`localStorage` の `blancPass` に保存して毎回送る
- 違うと「認証に失敗しました」→ アプリは保存を消して入力画面に戻す
- 4桁なので、間違いが 10 分に 30 回を超えたら 10 分間すべて拒否する（CacheService の `authFail`）
- GET は疎通確認（ping）だけ。JSONP は廃止（パスワードが URL に残るため）
- 変えるときはスクリプトプロパティを書き換えるだけ（各端末で入力し直しになる）

---

## メモ

- Terada（保管タブ）の一覧はセラー管理の「寺田倉庫」シートが本体（どの端末でも同じ）。
  アプリで WineList*.xlsx を読み込むとシートを作り直す（メモ・セラー在庫・セラーID は寺田IDで引き継ぐ）。
  `localStorage` の控えは通信を待たずに描くためだけ。nameMaster・wineCategories も同様に端末の控え。
- 日付は GAS 側で `Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd')` 済み。
- `gas/Code.gs` はバッククォートと `//` コメントを使わない書き方で統一している
  （既存コードの流儀に合わせたもの）。`index.html` 側にこの制限はない。
