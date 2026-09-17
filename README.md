# Blanc - White Wine Cellar

白ワインセラーの在庫・購入・テイスティング管理アプリ。

- **フロントエンド** … GitHub Pages（このリポジトリの静的ファイル）
- **バックエンド** … Google Apps Script の Web API（`doPost`）
- **データ** … Google スプレッドシート `1J4kQGKk0SOjKcVbgmlDIspDRhVcOLZ9J_u-tN4MAWlY`

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
| `saveWine` | wine オブジェクト | ワイン保存 | ○ |
| `deleteWine` | id 文字列 | ワイン削除 | ○ |
| `adjustStock` | `{id, delta}` | 在庫増減 | ○ |
| `savePurchase` | purchase オブジェクト | 購入履歴保存 | ○ |
| `saveTasting` | tasting オブジェクト | テイスティング保存 | ○ |
| `deleteTasting` | id 文字列 | テイスティング削除 | ○ |
| `saveDrinking` | drinking オブジェクト | 飲酒履歴保存 | ○ |

`Code.gs` の `apiHandlers_()` に載っていない関数は実行されない。
関数を増やしたら `apiHandlers_()` とこの表の両方に追加すること。

`callClaudeAPI` は API に公開していない（フロントからは呼んでおらず、
公開すると誰でも API キーを消費できてしまうため）。

---

## 更新のしかた

- **フロントを直した** … `index.html` を編集して push。加えて `sw.js` の
  `CACHE_VERSION` を上げる（`blanc-v1` → `blanc-v2`）。上げ忘れると
  古いキャッシュが残って更新が反映されない。
- **GAS を直した** … エディタで編集後、「デプロイを管理」から
  **既存デプロイを新バージョンで更新**（新規デプロイを作らない）。

---

## セキュリティ

### アクセス権は「全員」しか選べない

GAS のアクセス権を「自分のみ」にすると、GAS は accounts.google.com へ
リダイレクトを返す。そこには CORS ヘッダーが無いため、ブラウザは
クロスオリジンの fetch を打ち切る。**この構成で「自分のみ」は原理的に成立しない。**
旧 HtmlService 版にあった Google ログインという認証層は、移行によって失われている。

### 合言葉による保護

代わりに `API_TOKEN` で書き込みを守っている。

- **読み取り**（`READ_ONLY_FNS`）… 合言葉なしで通る
- **書き込み**（saveWine / deleteWine / adjustStock 等）… 合言葉が一致しないと拒否

`gas/Code.gs` と `index.html` の `API_TOKEN` は**必ず同じ文字列**にすること。
片方だけ変えると書き込みが全部通らなくなる。

ただし `index.html` は公開されるので、**合言葉もソースを読めば分かる**。
防げるのは「URL をたまたま見つけた第三者」と「公開リポジトリから GAS の URL を
自動収集するプログラム」まで。本気で調べる相手は防げない。

それ以上の保護が要る場合、この構成では手がない。GitHub Pages をやめて
GAS ウェブアプリに戻す（＝スマホでの見やすさを諦める）以外の選択肢はない。

---

## メモ

- `localStorage` で持っている Terada データ・nameMaster・wineCategories は
  移行不要（そのまま動く）。ただし**保存先ドメインが変わる**ため、
  旧 GAS 版に溜めていた内容は引き継がれない。必要なら手で移すこと。
- 日付は GAS 側で `Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd')` 済み。
- `gas/Code.gs` はバッククォートと `//` コメントを使わない書き方で統一している
  （既存コードの流儀に合わせたもの）。`index.html` 側にこの制限はない。
