# Excel 計算式・分析ロジック

計算はすべてブック内の数式（値のハードコードなし）。列は名前定義（例 `PH_Notional`＝06 の Notional 列全体）で参照するため、
数式バーで追跡できる。使用関数は Excel 2016 以降の MAXIFS／MINIFS／TEXTJOIN まで（XLOOKUP・FILTER・LET・動的配列は不使用）。
配列が必要な数式は従来型の配列数式（CSE）で保存している。

## 1. 基準日（00_Control）

| 名前 | 式 |
|---|---|
| EvalDate | 手入力があればその値、なければ `TODAY()-MOD(WEEKDAY(TODAY(),2)-5,7)`（直近金曜。月曜実行なら前週金曜） |
| PrevEvalDate | EvalDate−7 |
| MonthBaseDate / YearBaseDate | `EOMONTH(EvalDate,-1)` / 前年 12/31 |
| RunDate | 手入力（再現・監査用）またはTODAY() |

## 2. ポジション標準化（06_Position_History U 列以降）

| 項目 | ロジック |
|---|---|
| Index ID | Index Name を 02 B 表で照合。未登録は `UNMAPPED`（DQ09） |
| Underlying | Ticker の先頭語（`NDX US 11/20/26 P19500 Index` → NDX） |
| Option Type | Ticker の日付の次の語が `P`＋数字→PUT、`C`＋数字→CALL。だめなら Name に単語 `PUT`/`CALL`。どちらもなければ **UNKNOWN** |
| Position Key | 集約＝`Index ID｜Ticker`。個別取引＝集約キー｜open_date｜連番 |
| Days to Expiry | 満期日−基準日（暦日）、営業日＝`NETWORKDAYS(基準日,満期日,Holidays)-1` |
| Strike Distance | `(Spot−Strike)/Spot` |
| Moneyness | \|Dist\|≤ATM 幅（`S_AtmBand`）→ATM。PUT：Dist>0→OTM、<0→ITM。CALL：Dist<0→OTM、>0→ITM。判定不能→UNKNOWN |
| Snapshot | 指数ごとに、EvalDate 以前の最新基準日＝CUR、その前の基準日＝PRV、それ以外＝OLD |
| Position Status | CUR かつ満期日≥EvalDate→`CUR`（有効）。満期日<EvalDate→`CUR_EXPIRED`（**有効ポジションとして扱わない**。履歴には残る）。PRV も同様 |

Long/Short は Moneyness には影響させない（オプション自体の状態）。売り・買いの影響は Notional/Weight の符号で別途確認。

## 3. 指数単位の集計（10_Option_Analytics A 表）

対象は Status=CUR の行のみ。

| 指標 | 式 |
|---|---|
| ポジション数・Gross/Net Notional | COUNTIFS／SUMIFS(\|N\|)／SUMIFS(N) |
| Σ\|Weight\|・最大構成比 | SUMIFS(\|W\|)、MAXIFS(\|W\|)/Σ\|W\|（合計 100% は仮定しない） |
| Nearest Strike Distance | MINIFS(\|Dist\|)。該当銘柄・符号付距離・Moneyness も表示 |
| 加重平均 Strike Distance・IV | Σ(x×\|N\|)/Σ\|N\| |
| Nearest Maturity | MINIFS(満期)。営業日数は EvalDate から |
| ITM 数・Strike 5% 以内数 | COUNTIFS |
| Greeks | Σ(Greek×Notional) を**参考値**として表示（単位確認前）。Γ 変化率＝(今回−前回)/\|前回\| |
| 基準日遅れ | `NETWORKDAYS(ポジション基準日, EvalDate)-1` |
| Level vs BBG | ファイルの Index Level ÷ Bloomberg 指数値 −1（基準日一致時のみ。DQ19） |

## 4. ポジション比較（10 B 表）

集約キーごとに前回（PRV）・今回（CUR）の Notional 合計と行数を比較。

| 分類 | 条件 |
|---|---|
| New | 前回なし・今回あり |
| Expired | 前回あり・今回なし・満期日 ≤ 今回基準日 |
| Closed | 前回あり・今回なし・満期日 > 今回基準日 |
| Increased / Reduced | \|今回\|/\|前回\|−1 が ±`S_NotionalTol` 超 |
| Unchanged | 上記以外 |
| Strike Change | 同一キーで Strike 値が変化（Ticker が同じなのに Strike が違う＝要確認） |
| **Possible Roll（推定）** | 同一 Index・同一原資産・同一 Put/Call のグループで、New の満期より前の満期の Closed/Expired がある（新規側）、またはその逆（旧側）。**確定ではない**旨を表示 |
| ITM 変化 | 前回 Moneyness≠ITM かつ今回 ITM → `→ITM` |

## 5. NAV 照合（08 C 表）

比較前に「同一 Fund／同一基準日／同一通貨／同一 Basis」を確認する。

| 結果 | 条件 |
|---|---|
| MISSING DATA | 評価基準日以前の公式 NAV または Bloomberg NAV がない |
| NOT COMPARABLE | 通貨または Basis（TR/PRICE）が異なる（**差額アラートは出さない**） |
| DATE DIFFERENCE | 公式 NAV と同じ日付の Bloomberg NAV がない（**差額アラートは出さない**） |
| MATCH / WARNING / MISMATCH | \|BBG/公式−1\| ≤ `S_NavMatchTol` / ≤ `S_NavWarnTol` / 超過 |

NAV 公表遅延：評価基準日の公式 NAV が未取得で、`RunDate > WORKDAY(EvalDate, 公表ラグ)`。

## 6. リターン（08・09）

- 期間 TR：`(NAV_t + 期間内分配金) / NAV_{t-1} − 1`（分配金は Ex-Date が (t−1, t] のもの）。
- TRI：初回=100 として連鎖。**すべてのリターンは TRI の比率**（分配型も一貫して TR）。
- Weekly：評価基準日の公式 NAV（`S_NavDateTol` 以内）と、前週基準日以前の直近公式 NAV がそろう場合のみ。**どちらかが欠ければ空欄**（前回値で代用しない）。
- Monthly / YTD：前月末・前年末以前の直近 NAV を基準。
- Since Inception：TRI/100−1（履歴開始日＝設定日なら「設定来」、異なれば「履歴開始来」と表示）。
- Since Client Entry：取得日以前の直近 TRI を基準（`04`）。TRI がなければ取得時 NAV との価格比。
- Maximum Drawdown / 現在 DD：TRI の累積高値比。
- Realized Vol：直近 `S_VolWeeks` 週の週次 TR の標本標準偏差 ×√52（観測 4 未満は空欄）。

## 7. 寄与度（09 B 表）

| レベル | 条件・内容 |
|---|---|
| Confirmed | 運用会社の正式データ（09 C 表、Level=CONFIRMED）が評価基準日にある。Residual＝週次 TR−Σ構成要素 |
| Estimated | Vol Target：前週 Exposure×原指数騰落率（原資産寄与の推定）。為替ヘッジなし円建て：ドル円騰落率 |
| Observed | 上記がない場合。参照指数の騰落率を「観測」として示すのみで**要因は断定しない** |

## 8. Vol Target・Trend/Carry（11・12）

- Vol Target：Exposure 変化率＝今回/前回−1、Exp/Max、RV 変化（pt）、期間内リバランス回数、乖離＝Strategy 騰落−原指数騰落。
  「RV 上昇かつ Exposure 縮小」のとき【観測】として仕組みと整合的である旨を表示。
- Trend/Carry：銘柄ごとにシグナル符号反転・Long/Short 反転・Carry 変化。指数単位で Gross/Net、Long/Short 数、
  \|W\| 加重 Carry、最大構成比、アセットクラス別 Net。

## 9. アラート・優先度・コメント・メール

- アラート：docs/05。13 C 表で Mapping 行 × ルールを評価 → 発生分を重要度順に A 表へ（`LARGE` による並べ替え）。
- 営業優先度：docs/05。手動 Override 可。
- コメント（15）：定型文を数式で生成し、【事実】【推定】【観測】【示唆】【予定】を付ける。数値には基準日を付け、欠損時は【データ欠損】と明記。
  顧客向け Q&A には「（社内）」で始まる質問を使わない。Override 列があれば Final に優先採用。
- メール（16）：選択した担当者の保有行を優先度順に最大 10 件。注目ポイントはファンドごとに上位 3 アラート。
  DQ 未解決時は冒頭に【要確認】。末尾に社内限り・転送不可の注記。
