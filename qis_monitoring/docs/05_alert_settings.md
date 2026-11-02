# リスクアラート設定表

すべて `19_Settings` で変更可能（A 表＝閾値パラメータ、B 表＝ルール）。初期値は仮の値であり、商品特性に合わせて調整すること。
判定は `13_Risk_Alerts` の C 表（Fund-Index Mapping 行 × ルール）で行い、発生したものだけを A 表に重要度順で並べる。

- **Scope=FUND**：ファンド単位（各ファンドの最初の Mapping 行でのみ評価し、重複させない）
- **Scope=INDEX**：参照指数単位（1 ファンドが複数指数を持つ場合は指数ごとに評価）
- **Suppressed By**：上位ルールが発生している場合は表示しない（例：Strike 3% 以内が出ていれば 5% 以内は出さない）
- 説明文の `<v>`＝値、`<t>`＝閾値、`<x>`＝詳細（銘柄名など）に置換

## ルール一覧

| ID | Alert Type | Scope | Strategy | 指標 | 条件 | 閾値（初期値） | 重要度 | 抑制 | Recommended Sales Action | Data Source |
|---|---|---|---|---|---|---|---|---|---|---|
| R01 | 週次NAV下落 | FUND | ALL | `WEEKLY_RET` | <= | WeeklyNavDecline = -0.02 | High | — | パフォーマンス説明（下落要因は確認済み情報と推測を区別して説明） | 08_NAV_History（運用会社公式NAV） |
| R02 | 月次NAV下落 | FUND | ALL | `MONTHLY_RET` | <= | MonthlyNavDecline = -0.05 | High | — | 月次での下落局面説明と今後の見通し・リスク要因の共有 | 08_NAV_History |
| R03 | ドローダウン拡大 | FUND | ALL | `CUR_DD` | <= | Drawdown = -0.1 | Medium | — | ドローダウン状況と回復条件の説明 | 08_NAV_History（TR指数化） |
| R04 | NAV不一致 | FUND | ALL | `NAV_MISMATCH` | >= | 1 | High | — | 社内確認（顧客連絡前に運用会社・データ担当へ確認） | 08_NAV_History（公式 vs Bloomberg） |
| R05 | NAV照合 警告 | FUND | ALL | `NAV_WARNING` | >= | 1 | Low | — | データ担当による確認 | 08_NAV_History |
| R06 | NAV照合不能 | FUND | ALL | `NAV_UNCOMP` | >= | 1 | Medium | — | データ取得状況の確認 | 08_NAV_History |
| R07 | NAV公表遅延 | FUND | ALL | `NAV_DELAY` | >= | 1 | Medium | — | 運用会社へNAV公表状況を確認 | 08_NAV_History／02_Fund_Master（公表ラグ） |
| R08 | 異常リターン（要確認） | FUND | ALL | `ABN_RET` | ABS>= | AbnReturn = 0.1 | Medium | — | データ確認後に顧客対応を判断 | 09_Performance |
| R10 | Strike接近（3%以内） | INDEX | OPTION | `STRIKE_DIST` | <= | NearStrike2 = 0.03 | High | — | リスク状況確認・原資産下落時の挙動説明 | 06_Position_History（運用会社ポジションファイル） |
| R11 | Strike接近（5%以内） | INDEX | OPTION | `STRIKE_DIST` | <= | NearStrike1 = 0.05 | Medium | R10 | リスク状況確認 | 06_Position_History |
| R12 | ITMへの変化 | INDEX | OPTION | `ITM_TRANS` | >= | 1 | High | — | リスク状況確認・損益影響の確認（運用会社照会） | 06_Position_History |
| R13 | 満期接近（営業日） | INDEX | OPTION | `MAT_BDAYS` | <= | MatBizDays = 3 | Medium | — | 満期・ロール予定の確認 | 06_Position_History |
| R14 | 満期接近（30日以内） | INDEX | OPTION | `MAT_DAYS` | <= | MatDays = 30 | Low | R13 | 満期スケジュールの把握 | 06_Position_History |
| R15 | IV急上昇 | INDEX | OPTION | `IV_CHG` | >= | IvJump = 0.03 | Medium | — | ボラティリティ環境の説明 | 06_Position_History（ファイル上のImplied Vol） |
| R16 | Gamma急変 | INDEX | OPTION | `GAMMA_CHG` | ABS>= | GammaChg = 0.5 | Low | — | 感応度変化の把握（社内） | 06_Position_History |
| R17 | Weight集中度上昇 | INDEX | OPTION | `CONC_CHG` | >= | ConcUp = 0.1 | Low | — | ポジション構成の把握 | 06_Position_History |
| R18 | オプションロール（推定） | INDEX | OPTION | `ROLL_CNT` | >= | 1 | Low | — | ポジション変化の説明（ロールは推定である旨を明記） | 06_Position_History（前回比較） |
| R19 | ポジションデータ遅延 | INDEX | OPTION | `POS_STALE` | > | PosDelayBizDays = 3 | Medium | — | 最新ポジションファイルの受領確認 | 05_Position_Raw／06_Position_History |
| R20 | Exposure急縮小 | INDEX | VOLTARGET | `EXP_CHG` | <= | ExposureDrop = -0.2 | High | — | 戦略動向説明（Vol Targetの仕組み） | 11_VolTarget_Analytics（運用会社週次データ） |
| R21 | Exposure急拡大 | INDEX | VOLTARGET | `EXP_CHG` | >= | ExposureChg = 0.2 | Low | — | 戦略動向説明 | 11_VolTarget_Analytics |
| R22 | レバレッジ上限接近 | INDEX | VOLTARGET | `EXP_RATIO` | >= | LevCapRatio = 0.95 | Medium | — | 上限到達時の挙動説明 | 11_VolTarget_Analytics |
| R23 | 実現Vol急上昇 | INDEX | VOLTARGET | `RV_CHG` | >= | RvJump = 0.03 | Medium | — | 市場環境説明 | 11_VolTarget_Analytics |
| R24 | リバランス発生 | INDEX | VOLTARGET | `REBAL` | >= | 1 | Low | — | 戦略動向の把握 | 11_VolTarget_Analytics |
| R25 | 原指数との乖離 | INDEX | VOLTARGET | `TRACK` | ABS>= | TrackDiff = 0.02 | Low | — | 乖離要因の説明 | 11_VolTarget_Analytics |
| R26 | Trendシグナル反転 | INDEX | TRENDCARRY | `SIG_REV` | >= | 1 | Medium | — | 戦略動向説明（ポジション転換） | 12_TrendCarry_Analytics（運用会社週次ポジション） |
| R27 | Long/Short反転 | INDEX | TRENDCARRY | `LS_REV` | >= | 1 | Medium | — | 戦略動向説明 | 12_TrendCarry_Analytics |
| R28 | Carry低下 | INDEX | TRENDCARRY | `CARRY_CHG` | <= | CarryDrop = -0.01 | Low | — | 収益源の変化説明 | 12_TrendCarry_Analytics |
| R29 | ポジション集中 | INDEX | TRENDCARRY | `TC_CONC` | >= | PosConc = 0.4 | Low | — | 集中リスクの説明 | 12_TrendCarry_Analytics |
| R30 | 市場データ欠損 | INDEX | ALL | `MKT_MISSING` | >= | 1 | Medium | — | データ取得の確認（前回値で代用しない） | 07_Bloomberg_Data |

## 閾値・スコアのパラメータ（19_Settings A 表）

| Key | 初期値 | 単位 | 説明 |
|---|---|---|---|
| `S_DataMode` | MANUAL | MANUAL/BLOOMBERG | 市場データの採用元。BLOOMBERG＝BDH 取得値のみ採用／MANUAL＝手入力・CSV 値を採用（サンプル配布時は MANUAL） |
| `S_NavMatchTol` | 0.0005 | 比率 | NAV 照合：差が この値以内なら MATCH |
| `S_NavWarnTol` | 0.003 | 比率 | NAV 照合：差が この値以内なら WARNING、超えると MISMATCH |
| `S_NavDateTol` | 0 | 日 | 評価基準日と NAV 日付の許容差（0＝同一日のみ） |
| `S_AtmBand` | 0.005 | 比率 | ATM 判定の許容幅（\|Strike Distance\| ≤ この値なら ATM） |
| `S_NearStrike1` | 0.05 | 比率 | Strike 接近（Medium） |
| `S_NearStrike2` | 0.03 | 比率 | Strike 接近（High） |
| `S_MatBizDays` | 3 | 営業日 | 満期接近（営業日） |
| `S_MatDays` | 30 | 日 | 満期接近（暦日） |
| `S_IvJump` | 0.03 | Vol（小数） | 加重平均 IV の前回比上昇幅（0.03＝3pt） |
| `S_GammaChg` | 0.5 | 比率 | Gamma×Notional 合計の前回比変化率（絶対値） |
| `S_ConcUp` | 0.1 | 比率 | 最大 Weight 構成比の上昇幅（0.10＝10pt） |
| `S_NotionalTol` | 0.001 | 比率 | Notional 変化がこの比率以内なら Unchanged |
| `S_WeeklyNavDecline` | -0.02 | 比率 | 週次 NAV 下落 |
| `S_MonthlyNavDecline` | -0.05 | 比率 | 月次 NAV 下落 |
| `S_Drawdown` | -0.1 | 比率 | 現在ドローダウン |
| `S_ExposureDrop` | -0.2 | 比率 | Vol Target Exposure 急縮小（前週比） |
| `S_ExposureChg` | 0.2 | 比率 | Vol Target Exposure 急拡大（前週比） |
| `S_LevCapRatio` | 0.95 | 比率 | Exposure ÷ 上限 がこの値以上で上限接近 |
| `S_RvJump` | 0.03 | Vol（小数） | 実現 Vol 前週比上昇幅 |
| `S_TrackDiff` | 0.02 | 比率 | Strategy と原指数の週次騰落率差（絶対値） |
| `S_CarryDrop` | -0.01 | 比率 | 加重平均 Carry の前週比低下幅（-0.01＝-1pt） |
| `S_PosConc` | 0.4 | 比率 | Trend/Carry 単一銘柄の \|Weight\| 構成比 |
| `S_PosDelayBizDays` | 3 | 営業日 | ポジション基準日が評価基準日よりこの営業日数を超えて古い場合に遅延 |
| `S_MktStaleDays` | 0 | 日 | 市場データ As-of が評価基準日よりこの日数を超えて古い場合 STALE |
| `S_AbnReturn` | 0.1 | 比率 | 週次リターンの絶対値がこの値以上なら異常値として確認 |
| `S_VolWeeks` | 26 | 週 | 実現ボラティリティの計測期間 |
| `S_IdxLvlTol` | 0.001 | 比率 | ポジションファイルの Index Level と Bloomberg 指数値の許容差 |
| `S_SevScoreHigh` | 3 | 点 | 営業優先度：High アラート 1 件あたり |
| `S_SevScoreMed` | 2 | 点 | 営業優先度：Medium アラート 1 件あたり |
| `S_SevScoreLow` | 1 | 点 | 営業優先度：Low アラート 1 件あたり |
| `S_LargeBalance` | 1000000000 | 円 | 残高大（+1 点） |
| `S_RedemptionDays` | 90 | 日 | 償還・満期予定がこの日数以内（+2 点） |
| `S_ContactOverdueDays` | 30 | 日 | 前回コンタクトからの経過日数超過（+1 点） |
| `S_FollowupWindowDays` | 7 | 日 | 次回フォロー予定がこの日数以内（+1 点） |
| `S_SevCap` | 2 | 件 | 重要度ごとに加点するアラート件数の上限（同一事象の指数別重複による過大評価を防ぐ） |
| `S_PriorityHigh` | 8 | 点 | 合計点がこの値以上で High |
| `S_PriorityMed` | 4 | 点 | 合計点がこの値以上で Medium |
| `S_ReportPrefix` | 【QIS Weekly Update】担当先ファンド状況｜ | 文字列 | メール件名の接頭辞 |
| `S_WorkbookVersion` | v1.0.0 | 文字列 | ブックのバージョン（メール履歴に記録） |

## 依頼仕様の初期閾値との対応

| 依頼仕様 | 実装 |
|---|---|
| Weekly NAV Decline -2% | R01（S_WeeklyNavDecline） |
| Monthly NAV Decline -5% | R02（S_MonthlyNavDecline） |
| Drawdown -10% | R03（TR 指数の高値からの現在ドローダウン） |
| Option Strike Distance 5% 以内 / 3% 以内 | R11（Medium）/ R10（High） |
| Maturity 30 日以内 / 3 営業日以内 | R14（Low）/ R13（Medium） |
| Exposure Change 前週比 20% 以上 | R20（縮小 High）/ R21（拡大 Low） |
| NAV Delay 公表予定日超過 | R07（評価基準日＋公表ラグ営業日 < 作成日 かつ 基準日 NAV 未取得） |
| Position Data Delay 更新予定日超過 | R19（ポジション基準日が評価基準日より 3 営業日超古い） |

## 営業優先度スコア（14_Sales_Priority）

`Score = min(High件数,S_SevCap)×S_SevScoreHigh + min(Medium件数,S_SevCap)×S_SevScoreMed + min(Low件数,S_SevCap)×S_SevScoreLow`
`+ 1（残高 ≥ S_LargeBalance） + 2（償還・満期まで ≤ S_RedemptionDays 日） + 1（前回接触から > S_ContactOverdueDays 日、または記録なし） + 1（次回フォロー予定が作成日から S_FollowupWindowDays 日以内）`

Score ≥ S_PriorityHigh → High、≥ S_PriorityMed → Medium、それ以外 Low。Override 列に入力すると手動値が優先され、Reason に「（手動変更）」と表示される。

## データ品質チェック（18_Data_Quality）

Critical が 1 件でも OPEN の場合、`DQ_Overall`＝要確認 となり、メール本文冒頭に【要確認】、Outlook 下書きの件名にも【要確認】が付く。確認のうえ問題なしと判断した場合は Ack 欄に Y と確認者・確認内容を記入する。

| ID | チェック | 区分 | 重大度 |
|---|---|---|---|
| DQ01 | 評価基準日の公式NAV欠損（ACTIVE） | NAV | Critical |
| DQ02 | NAV照合 MISMATCH | NAV | Critical |
| DQ03 | NAV照合 WARNING／照合不能 | NAV | Warning |
| DQ04 | NAV通貨がマスタと不一致（Currency Mismatch） | NAV | Critical |
| DQ05 | NAV履歴の重複キー | NAV | Critical |
| DQ06 | 異常リターン（Abnormal Return） | NAV | Warning |
| DQ07 | ポジション基準日の遅延（オプション指数） | Position | Critical |
| DQ08 | ポジション基準日 ≠ 評価基準日（時点不一致） | Position | Warning |
| DQ09 | Index Name 未登録（Fund-Index マッピング不能） | Position | Critical |
| DQ10 | Fund-Index Mapping 不整合／Primary 未設定 | Master | Critical |
| DQ11 | ポジション行の重複（Position ID Duplicate） | Position | Critical |
| DQ12 | Ticker 欠損（Missing Ticker） | Position | Critical |
| DQ13 | Put/Call 判定不能（UNKNOWN） | Position | Warning |
| DQ14 | 最新ファイル内の満期経過ポジション | Position | Warning |
| DQ15 | 同一基準日・同一指数で複数ファイル（再取込の衝突） | Position | Warning |
| DQ16 | 重要市場データの欠損・STALE | Market | Critical |
| DQ17 | その他市場データの欠損・STALE | Market | Warning |
| DQ18 | 指数マスタの市場データKeyが07に未設定 | Market | Warning |
| DQ19 | Index Level（ファイル）と Bloomberg 指数値の乖離 | Market | Warning |
| DQ20 | ACTIVE ファンドに保有先なし（Missing Client Holdings） | Client | Warning |
| DQ21 | 保有データの顧客/ファンド未登録 | Client | Critical |
| DQ22 | Greeks の単位・符号が未確認（参考値扱い） | Position | Warning |
| DQ23 | NDX Spot 列の有効性未確認（原資産が NDX 以外の可能性） | Position | Warning |
| DQ24 | 市場データが MANUAL モード（Bloomberg 未使用） | Market | Warning |
