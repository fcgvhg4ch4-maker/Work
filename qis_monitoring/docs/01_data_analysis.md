# ポジションファイル構造分析と Phase 1 の必要データ

## 0. 前提（重要）

**依頼文で「添付」とされたポジションファイルは、このリポジトリ・作業環境には含まれていませんでした。**
このため、構造分析は追加仕様 A に列挙された 17 列の定義に基づいて行い、ブックの検証には同じ 17 列を持つ
**架空のサンプルファイル**（`sample_data/positions/positions_20261023.csv`・`positions_20261030.csv`）を使っています。
実ファイルを受領したら、以下の「要確認」事項を最初に潰してください。列名が異なる場合は `19_Settings` の
`tblColMap`（Power Query 列マッピング）の SourceColumn を書き換えるだけで取り込めます。

## 1. ファイル構造（追加仕様 A に基づく想定）

| 項目 | 想定 | 確度 | 要確認事項 |
|---|---|---|---|
| シート／テーブル | 1 ファイル 1 表。複数 Index Name が縦に連結 | 推定 | xlsx の場合のシート名（`tblPQParams.PositionSheetName`）。見出し行の位置、注記行・合計行の有無 |
| 形式 | CSV（UTF-8）または xlsx | 推定 | 文字コード（Shift_JIS の場合は `fnReadPositionFile` の Encoding を 932 に） |
| 1 行の粒度 | 1 取引（同一 Ticker が open_date 別に複数行あり得る） | 推定 | 同一 Ticker 複数行が別取引か重複送付か。Trade ID の有無 |

## 2. 列の分析（依頼文 4 章の 12 項目）

| # | 確認項目 | 該当列 | 判断 | 要確認事項 |
|---|---|---|---|---|
| 1 | シート・テーブル構造 | — | 上表 | — |
| 2 | ファンドを識別する列 | **なし**（`Index Name` は参照 QIS 指数） | ファイルにファンド ID はない。ファンドとの対応は `02_Fund_Master` の Fund-Index Mapping（1 対多）で管理 | 1 指数を複数ファンドが参照するか、配分比率 |
| 3 | ポジションを識別する列 | `BloombergTicker`（＋`Index Name`、`open_date`） | 集約キー＝Index ID｜Ticker、取引キー＝集約キー｜open_date｜連番 | Trade ID の提供可否 |
| 4 | 基準日 | `Position Date` | 1 ファイル内で指数ごとに単一であることを検証 | 時点（NY 終値か東京時間か） |
| 5 | 原資産・参照指数・Ticker | `Index Name`／`BloombergTicker`（原資産は Ticker 先頭語）／`NDX Spot` | Ticker 例 `NDX US 11/20/26 P19500 Index` から原資産・Put/Call を解析 | 原資産 Bloomberg 識別子（NDX Index）と QIS 指数自体の Ticker |
| 6 | 数量・Notional・MV | `Notional(USD)`、`Weight`、`Option Mid Price` | 数量列・時価列はない。MV は Mid×数量で求められない（数量・乗数が不明） | Notional の符号（売り＝負か）、単位（額面か指数単位か）、Weight の定義 |
| 7 | 通貨 | 列名 `Notional(USD)` から USD | 通貨列はない。指数通貨は `02` の Index マスタで管理 | USD 以外の指数が来るか |
| 8 | Strike・Barrier・Maturity | `Strike`、`Maturity Date` | **Barrier 列はない** → Barrier 分析は未実装（データ待ち） | Barrier 型商品の有無 |
| 9 | Option Type・Greeks | Type 列なし（Ticker／Name から解析）、`Delta` `Gamma` `Vega` `Theta` `Implied Vol` | 判定不能は UNKNOWN（`18` DQ13 で検出） | Greeks が 1 単位あたりか、ポジション全体か。IV の単位（0.25 か 25 か） |
| 10 | Vol Target・Leverage・Exposure | **なし** | `11` の入力表（運用会社の週次データ）で補う | 運用会社から取得できる項目と頻度 |
| 11 | Trend／Carry シグナル | **なし** | `12` の入力表で補う | 同上 |
| 12 | 前週比較に必要な情報 | `Position Date`＋Ticker＋open_date＋Notional | 前回スナップショットを `06` に値で保存し、集約キーで比較 | — |

### 列ごとの推定と確認事項（列名が不明確なもの）

| 列 | 推定した意味（ブックでの扱い） | 確認が必要な点 |
|---|---|---|
| `Name` | Bloomberg 銘柄名。Ticker で判定できない場合の Put/Call 補助判定に使用 | 表記ゆれ |
| `Option Mid Price` | 基準日の Mid。Bloomberg PX_MID と照合（基準日一致時のみ） | 価格単位（指数ポイント）・乗数 |
| `Notional(USD)` | 絶対値で Gross、符号付き合計で Net を計算 | 符号・単位 |
| `Delta`〜`Theta` | **1 単位あたり**と仮定し Σ(Greek×Notional) を**参考値**として表示 | 単位・符号（確認後に `02` で「Greeks 単位確認済」＝Y） |
| `Implied Vol` | 小数（0.248＝24.8%） | 単位 |
| `NDX Spot` | 当該指数の原資産スポット（NDX）として Strike Distance に使用 | NDX 以外の原資産の指数でもこの列が NDX なのか |
| `Index Level` | 基準日の QIS 指数水準。前回比を指数騰落率として表示し、Bloomberg 指数値と照合（DQ19） | 算出時点 |
| `Weight` | 符号付き構成比。絶対値で構成比を計算（合計 100% は仮定しない） | 分母（NAV か名目か） |
| `open_date` | 取引開始日。同一 Ticker の別取引を区別 | 約定日か受渡日か |

## 3. Phase 1（MVP）に必要なデータと不足項目

| データ | 用途 | 取得元 | 状況 |
|---|---|---|---|
| ポジションファイル（週次） | ポジション取込・指数単位分析 | 運用会社 | 列定義は確定。**実ファイル未受領** |
| ファンドマスタ（Fund ID・通貨・設定日・NAV 公表ラグ・為替ヘッジ・分配有無） | NAV 管理・リターン計算 | 社内／運用会社 | **要作成**（サンプル値で代替中） |
| Fund-Index Mapping | ファンドと Index Name の対応 | 運用会社・商品部 | **要作成** |
| 公式 NAV 履歴（設定来または前年末以降） | 週次・月次・YTD・設定来・DD・Vol | 運用会社 | **要受領**（形式未定） |
| 分配金（Ex-Date・金額） | TR 計算 | 運用会社 | 分配型ファンドのみ **要受領** |
| Bloomberg NAV Ticker | NAV 照合 | Bloomberg | **要確認** |
| 原資産・QIS 指数・市場データ Ticker | 市場コメント・指数騰落率 | Bloomberg | 主要指数は一般的な Ticker を設定済み。**QIS 指数 Ticker は架空** |
| 地銀・担当者・保有（口数・簿価・取得日） | 営業優先度・メール | 社内 | **要作成**（顧客情報のため社内管理） |
| 前回コンタクト日・次回フォロー予定 | 営業優先度 | 社内（CRM 等） | 手入力欄あり |
| Vol Target／Trend・Carry の週次データ | 戦略別分析（Phase 2） | 運用会社 | ポジションファイルに含まれない。**要依頼** |
| 正式な寄与度データ | Performance Attribution（Confirmed） | 運用会社 | **要依頼**（なければ Observed 表示） |
| Barrier 情報 | Barrier 接近・発生の区別 | 運用会社 | **列なし**（Barrier 分析は保留） |

## 4. 実データ構造に応じた最適化（実施内容）

- 依頼の 21 シート構成は維持し、オプション中心のポジションファイルに合わせて次の点を最適化した。
  - ファンドではなく **Index Name 単位**で集計（`10_Option_Analytics` A 表）し、Mapping でファンドに展開。
  - ポジション比較は **集約キー（Index｜Ticker）** で行い、同一 Ticker の複数取引は合算（個別取引は `10` C 表で確認）。
  - Barrier・Leverage・Signal 等ファイルにない項目は、専用の入力表（`11`・`12`）として分離し、「Requires Additional Data」と明示。
- 追加シート：`21_Impl_Status`（実装区分・本番前確認事項）、`22_Test_Results`（テスト結果）。
