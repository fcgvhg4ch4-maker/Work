# Bloomberg フィールドマッピング

ブックでは `07_Bloomberg_Data` に集約する。**開発環境には Bloomberg 端末がないため、フィールド名・Ticker・単位は
すべて端末の `FLDS`／`SECF` で確認が必要**（ブック上の「確認状況」列＝確認要）。

## 1. 取得方式

| 関数 | 用途 | ブックでの使い方 |
|---|---|---|
| **BDH** | 時系列・特定日の値 | 評価基準日の値は `BDH(Ticker, Field, EvalDate, EvalDate, "Dates=H", "Fill=N")`。値がなければ空＝MISSING（前日値で埋めない）。前週・前月末は `Fill=P`（直前営業日の終値を基準値として使用） |
| **BDP** | 最新値（ライブ） | 接続テスト・オプション Greeks/IV の参考照合のみ。**レポートの数値には使わない**（時点が評価基準日と異なるため） |
| **BDS** | バルクデータ | 分配金履歴（DVD_HIST_ALL）など。現状は未使用（Future Enhancement） |

## 2. 採用ルール（古い値を最新値として扱わない）

`19_Settings` の `S_DataMode`：

- **BLOOMBERG**：BDH 列の数値のみ採用。取得できなければ空欄（`Status=MISSING`）。手入力値は使わない。
- **MANUAL**：手入力・CSV 列（K〜N）を採用。`As-of` が評価基準日より `S_MktStaleDays` 日を超えて古い場合は `STALE` とし、**検証済値（U 列）には入れない**。

下流（参照指数騰落率・市場コメント・アラート）はすべて検証済値のみを参照する。欠損・STALE は 18 の DQ16（重要）・DQ17 と、
13 の「市場データ欠損」アラートで表示する。

## 3. マーケットデータ一覧（初期設定）

| Key | 内容 | Ticker | Field | 単位 | 重要 | 備考 |
|---|---|---|---|---|---|---|
| EQ_NDX | NASDAQ 100 | NDX Index | PX_LAST | PRICE | Y | オプション指数の原資産 |
| EQ_SPX | S&P 500 | SPX Index | PX_LAST | PRICE | Y | Vol Target 原指数 |
| EQ_SX5E | EURO STOXX 50 | SX5E Index | PX_LAST | PRICE | N | |
| EQ_NKY | 日経平均 | NKY Index | PX_LAST | PRICE | N | |
| FUT_NQ | NASDAQ100 先物 | NQ1 Index | PX_LAST | PRICE | N | ジェネリック |
| VOL_VIX | VIX | VIX Index | PX_LAST | VOL | Y | |
| VOL_VXN | VXN | VXN Index | PX_LAST | VOL | Y | NDX のボラティリティ指数 |
| RV_NDX | NDX 30 日実現 Vol | NDX Index | VOLATILITY_30D | VOL | N | フィールド要確認 |
| IV_NDX | NDX 1M ATM IV | NDX Index | 30DAY_IMPVOL_100.0%MNY_DF | VOL | N | フィールド要確認 |
| RT_UST2/5/10 | 米国債 2/5/10 年 | USGG2YR / USGG5YR / USGG10YR Index | PX_LAST | YIELD | N | 変化は bp 表示 |
| RT_JGB2/5/10 | 国債 2/5/10 年 | GJGB2 / GJGB5 / GJGB10 Index | PX_LAST | YIELD | N | |
| SW_USD10 | USD SOFR OIS 10 年 | USOSFR10 Curncy | PX_LAST | YIELD | N | Ticker 要確認 |
| SW_JPY10 | JPY TONA OIS 10 年 | JYSO10 Curncy | PX_LAST | YIELD | N | Ticker 要確認 |
| FX_USDJPY | ドル円 | USDJPY Curncy | PX_LAST | FX | Y | |
| FX_EURJPY | ユーロ円 | EURJPY Curncy | PX_LAST | FX | N | |
| FX_EURUSD | ユーロドル | EURUSD Curncy | PX_LAST | FX | N | |
| CR_CDXIG | CDX IG 5Y | CDX IG CDSI GEN 5Y Corp | PX_LAST | SPREAD | N | Ticker 要確認 |
| CMD_CL / CMD_GC | WTI・金（期近） | CL1 Comdty / GC1 Comdty | PX_LAST | PRICE | N | |
| QIS_* | QIS 指数 | **架空 Ticker（SQNDXPW Index 等）** | PX_LAST | PRICE | Y | 運用会社・指数算出者に確認 |

指数ごとの原資産 Key・QIS 指数 Key・Vol Key は `02_Fund_Master` B 表で設定（例：NDX 関連指数 → `EQ_NDX`／`VOL_VXN`）。

## 4. フィールド辞書（07 の B 表と同内容）

| データ項目 | Field | 関数 | 用途・注意 |
|---|---|---|---|
| 終値 | PX_LAST | BDH／BDP | 評価基準日は BDH（Fill=N） |
| ファンド NAV | FUND_NET_ASSET_VAL／PX_LAST | BDH | Ticker によりフィールドが異なる。照合用 |
| NAV 日付 | NAV_DATE／LAST_UPDATE_DT | BDP | 基準日の確認 |
| 分配金履歴 | DVD_HIST_ALL | BDS | 分配型の TR 調整（将来） |
| オプション Mid | PX_MID | BDH／BDP | ファイルの Option Mid Price と照合（基準日一致時のみ） |
| Greeks | OPT_DELTA_MID 等 | BDP | 過去時点は取得できない可能性 → 時点不一致として表示 |
| Implied Vol | IVOL_MID | BDP | 単位要確認 |
| オプション属性 | OPT_STRIKE_PX／OPT_EXPIRE_DT／OPT_PUT_CALL／OPT_UNDL_TICKER | BDP | Ticker 解析結果の検証 |
| 実現 Vol | VOLATILITY_30D／VOLATILITY_260D | BDP／BDH | 年率 % |
| ATM IV | 30DAY_IMPVOL_100.0%MNY_DF | BDP／BDH | |

## 5. オプション照合（07 D 表）

現在の有効ポジションの Ticker ごとに、ファイル値と Bloomberg 値を並べる。ポジション基準日と取得時点が異なる場合は
「時点不一致（Greeks/IV はライブ値）」と表示し、差を誤差として扱わない。

## 6. データ利用上の注意

- Bloomberg データの保存（07 の時系列表・値貼付け）・社内共有・メール記載は**契約条件に従う**。範囲は本番前に確認（docs/08）。
- **顧客向けの転送処理は実装していない。** メール本文末尾に「社内限り・顧客への転送不可」を自動で記載する。
