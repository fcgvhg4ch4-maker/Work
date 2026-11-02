# QIS Fund Monitoring & Weekly Sales Intelligence Workbook

地銀担当営業向けに、QIS ファンド（Option-Based／Volatility Target／Trend Following・Carry）の週次モニタリングと
営業メール本文を生成する Excel ワークブック。Excel 365＋Power Query＋Bloomberg Excel Add-in（＋必要時のみ VBA）を前提とする。

> **サンプルデータはすべて架空**（ファンド・指数・地銀・価格・Ticker の QIS 指数部分）。依頼文の「添付ポジションファイル」は
> 作業環境に含まれていなかったため、追加仕様 A の 17 列定義に基づく架空ファイルで構築・検証している（docs/01）。

## 納品物

| # | 納品物 | 場所 |
|---|---|---|
| 1 | 実務利用可能な Excel ワークブック（23 シート、数式ベース、架空サンプル入り） | `dist/QIS_Fund_Monitoring_Workbook.xlsx` |
| 2 | Bloomberg フィールドマッピング表 | `docs/02_bloomberg_field_mapping.md`／ブック `07_Bloomberg_Data` |
| 3 | Power Query 取り込み仕様＋M コード | `docs/03_power_query_spec.md`／`powerquery/*.pq` |
| 4 | Excel 計算式・分析ロジック | `docs/04_calculation_logic.md` |
| 5 | リスクアラート設定表 | `docs/05_alert_settings.md`／ブック `19_Settings` |
| 6 | 営業担当者別週次メール生成機能 | ブック `16_Email_Generator`＋`vba/modQIS.bas`（Outlook 下書き） |
| 7 | 操作マニュアル | `docs/06_operation_manual.md`／ブック `20_Instructions` |
| 8 | テスト結果（83 項目） | `docs/07_test_results.md`／ブック `22_Test_Results` |
| 9 | 未実装機能一覧 | `docs/08_gaps_and_preprod.md`／ブック `21_Impl_Status` |
| 10 | 本番導入前の確認事項 | 同上 |
| — | ポジションファイル構造分析・Phase 1 の必要データと不足項目 | `docs/01_data_analysis.md` |

## 実装区分（要約）

| 区分 | 内容 |
|---|---|
| **Fully Implemented**（Excel 数式） | ポジション標準化・前回比較（New/Closed/Expired/Increased/Reduced/Possible Roll/ITM 変化）、指数単位集計、NAV 照合（MATCH/WARNING/MISMATCH/DATE DIFFERENCE/MISSING DATA＋比較不能）、TR リターン・DD・Vol、Vol Target・Trend/Carry 分析、29 ルールのアラート、営業優先度、定型コメント、担当者別メール本文、データ品質 24 項目、ダッシュボード |
| **Requires Bloomberg Connection** | 07 の BDH/BDP 数式（未接続時は NO_CONN → MANUAL モードで手入力値を使用）。フィールド・Ticker は端末で要確認 |
| **Requires Local Excel Configuration** | Power Query（M コードを同梱。ブックには未埋め込み） |
| **Requires VBA / Outlook Permission** | 一括更新、ポジション履歴の値保存（重複防止）、Outlook 下書き作成（送信しない）、メール履歴保存 |
| **Requires Additional Data** | Vol Target・Trend/Carry の週次データ、正式寄与度、公式 NAV・分配金、Barrier 情報 |
| **Future Enhancement** | 高度な自然言語生成（承認済み AI 環境が前提）、BDS 分配金取得、Trade ID による確定ロール判定 など |

詳細は `docs/08_gaps_and_preprod.md`。

## シート構成

| 論理区分 | シート |
|---|---|
| Control / Output | 00_Control、01_Dashboard、15_Weekly_Commentary、16_Email_Generator |
| Master | 02_Fund_Master（ファンド・指数・Fund-Index Mapping）、03_Client_Master、04_Client_Holdings、19_Settings |
| Raw Data | 05_Position_Raw、06_Position_History、07_Bloomberg_Data、08_NAV_History（＋11・12・09 の入力表） |
| Analytics | 09_Performance、10_Option_Analytics、11_VolTarget_Analytics、12_TrendCarry_Analytics、13_Risk_Alerts、14_Sales_Priority |
| Audit | 17_Email_History、18_Data_Quality、20_Instructions、21_Impl_Status、22_Test_Results |

## 再生成・検証

```sh
cd build
./make.sh        # sample_data → build_workbook → verify_workbook（LibreOffice 再計算）→ 結果埋め込み → gen_docs
```

- `build/sample_data.py` — 架空データの正本（テストシナリオの定義を含む）。`sample_data/` に CSV を書き出す
- `build/build_workbook.py` — ブック生成（openpyxl）。容量（ファンド 30・保有 100 等）は冒頭の定数
- `build/verify_workbook.py` — LibreOffice で再計算し、Python で独立に計算した期待値と照合。入力を書き換えた派生シナリオ（重複取込・Bloomberg 未接続・NAV 欠損・DQ 承認・担当者切替・手動変更・閾値変更）も検証
- `build/gen_docs.py` — `docs/05`・`docs/08` をコードから生成

検証は LibreOffice 24.2 で実施。Excel 365 での初回オープン時に `22_Test_Results` の値と一致するか確認すること（docs/08）。

## 設計上の原則

- 公式 NAV を優先し、Bloomberg NAV は照合用。同一基準日・同一定義でなければ差額アラートを出さない。
- 欠損・古いデータを最新値として扱わない（前回値で代用しない）。欠損はコメント・メールにも明示。
- 正式な寄与度データがない限り NAV 変動要因を断定しない（Confirmed／Estimated／Observed、【事実】【推定】【観測】【示唆】）。
- Greeks の合計は単位確認まで「参考値」。ロール判定は「推定」。
- メールは下書きまで（自動送信しない）。Bloomberg 由来データを顧客向けに転送する処理は実装しない。
