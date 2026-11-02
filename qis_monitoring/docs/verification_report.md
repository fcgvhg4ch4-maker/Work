# 検証レポート — QIS Fund Monitoring & Sales Update Workbook

作成日：2026-11-02
対象：`dist/QIS_Fund_Monitoring.xlsx`（本番テンプレート）／`dist/QIS_Fund_Monitoring_TESTDATA.xlsx`（架空データ版）
規模：17 シート、Excel テーブル 21、名前定義 138、数式セル約 18 万（VBA・マクロなし）

## 1. 最重要の前提（必ずお読みください）

| 項目 | 状況 |
|---|---|
| 添付ファイル `end_investor_holdings_view.xlsx`・STEP ポジションデータ | **作業環境に存在せず、入手できていません。** 本番テンプレートの入力テーブルは空です。架空の値で補完していません。 |
| 「別途提示された元の仕様書」 | **入手できていません。** シートごとの項目・判定条件・閾値・表示要件は、本マスタープロンプトの記載から設計しました。元仕様書との照合が必要です。 |
| グレー列（元データ列）の列名・順番 | 実ファイルの列構成が不明なため、**仮の設計**です。実ファイルと照合のうえ列を合わせてください（操作説明 付録 B）。 |
| アラート閾値 | 元仕様書の値が不明なため、**暫定値**を入れています。`00_Control` C12（閾値照合済み）が `Yes` になるまで、READY は表示されません（99_Check C02）。 |
| 15_QA_Output の定義 | 元仕様書がないため「顧客想定 Q&A」と解釈して実装しました。品質保証（QA）出力の意図だった場合は修正が必要です。 |
| Excel 365 本体での動作 | この環境には Excel がないため、**LibreOffice で再計算して検証しました。** Excel 365 で開いたうえでの確認（特に下記 4. の項目）が未実施です。 |

## 2. 実装済み機能

| 区分 | 内容 |
|---|---|
| データモデル | Fund（fund_code）／Investor Holding（end_investor_bdr_id × fund_code）／Strategy（strategy_data_key）／Position（strategy_data_key × position_date × position_key）の 4 階層を分けて管理。NAV・AUM・リターンはファンド固有の時系列からのみ計算 |
| 01 Live Fund | fund_code の重複を検出して表示（削除しない）。同一 Index を参照するファンド数・保有投資家数・保有額 |
| 02 Holdings | 明細（元データ全行）と集約ビューを分離。Amount が非数値のキーは集約額を空欄にして Check Required（0 扱いしない）。Live にない fund_code も検出 |
| 03 Strategy Map | ①fund_code 明示 → ②Index 承認済 → ③手動（承認者必須）の順で決定。同一 Index に複数キーがある場合は自動選択せず Review Required。multiplier は根拠と確認が揃ったときだけ使用 |
| 04/07 NAV | 1D・1W（7 暦日前）・MTD・3M・YTD・1Y・SI。基準日以前で直近の有効 NAV を使用し、許容ギャップを超える期間・履歴が足りない期間は空欄。NAV 日の重複は両方を除外。現在 DD・最大 DD（累積最高値基準）。AUM 変化と推定フロー（公式フローとは区別して表示）。TR は定義が確認済みの場合のみ計算 |
| 05 Market | 参照指数リターン（ファンドと同じ基準日にそろえて計算）。通貨・ヘッジ未確認の指数は「Reference only」と表示し、差分は出さない。原資産 → Spot 系列の確認済み対応表 |
| 06/08 Position | 原本は改変しない。最新版の判定（received_datetime → file_seq）と旧版の保持。Composition を優先し Portfolio で補完・照合。Ticker 解析値は別列に保持し、原本値と矛盾する場合は Check/STOP。Spot と Strike の原資産が一致しない場合は計算を停止。Strike Distance（Strike/Spot−1）、Put/Call 別の ITM/OTM/ATM、Buffer の表示（絶対値に ITM/OTM を併記）、残日数。Reported Greeks と Weight-adjusted Greeks を区別（定義が不明なら計算しない）。Premium 4 種は合算しない |
| 09 Position Change | 同じキーの直近 2 有効スナップショットで比較。New / Existing / Matured / Removed。Weight・Delta の差分。非一意キーは Check Required。ロールは「Potential Roll（要確認）」の表示のみ |
| 10 Alert | 15 種 × ファンド。各アラートに実績値・閾値（Control 参照）・重要度・理由・データソース・評価日・判定日時・alert_id・確認状況・レビュー担当・解消状況を保持。データ欠損は独立したアラート。欠損がある場合は GREEN にならない |
| 11 Coverage | ファンド別の欠損項目、Complete/Partial/Missing、データゲート |
| 12 Distribution | 投資家別：保有ファンド数・一覧・Amount・更新対象・担当・宛先・CC・最終配信日・次回予定・配信ステータス・レビュー状況。個別配信と統合配信を選択可能。投資家×ファンド単位の配信判定と理由。メール送信は実装していない |
| 13 Dashboard | 1 投資家 × 1 ファンド。上段にヘッダー、中段にパフォーマンス・参照指数比較・ポジション（最大 12 レッグ）、下段にサマリー。起点=100 に正規化した NAV と参照指数の折れ線グラフ。A4 横で印刷できる設定 |
| 14 Comment | 7 セクションのテンプレート（件名〜データ注意事項）。欠損時は【データ不足】と表示。見通しは書かない。comment_id（評価日・NAV 基準日・Position 基準日を含む）単位で reviewed_text → Approved → final_text |
| 15 QA | 想定 Q&A 8 問（要因・見通しは回答しない設計） |
| 99 Check | 36 項目。Check ID / Category / Description / Status / Error Reason / Affected Fund / Required Action / Owner / Last Checked。全体ゲートと最終判定（READY / REVIEW REQUIRED / DATA REQUIRED） |
| デザイン | 入力＝淡い黄、計算＝淡い青、元データ＝淡いグレー。重要度と判定の色分けは条件付き書式。各シートの A1〜A5 に目的・データソース・更新方法・注意事項 |

## 3. 未実装・制約（実装済みとしては扱っていません）

1. **XLOOKUP / FILTER / UNIQUE / SORT / LET は使っていません。** 検証に使った LibreOffice 24.2 がこれらの関数を計算できないためです。同等の処理を INDEX/MATCH・MAXIFS・MINIFS・SUMIFS・COUNTIFS・TEXTJOIN・NETWORKDAYS.INTL で書いています。Excel 365 で動くことは変わりませんが、ご指示の関数構成とは異なります。
2. **集約ビューと計算ミラーは容量が固定です**（UNIQUE/FILTER のスピルを使っていないため）：ファンド 60、保有の集約 400、ポジション 1,500 行。容量を超えた場合は 99_Check（C06/C16/C17）で DATA REQUIRED を表示します。拡張するには `build_workbook.py` の CAP_* を変更して作り直してください。元データのテーブル（01/02/04/05/06 など）は行を追加すると自動で拡張されます。
3. **Power Query は未実装です。** 貼り付けで更新する構造です（操作説明 付録 A に手順方針のみ記載）。
4. **Excel 365 での未確認事項**：
   - テーブル内の単一セル配列数式（12 tbl_Dist の fund_list / update_target_funds）
   - Dashboard と 99_Check の配列数式（SMALL(IF()), TEXTJOIN(IF())）。Excel では `{}` 付きで表示される見込みです
   - 条件付き書式とグラフの描画
5. STEP 以外（Target Return / Target Buying / CTA / Rates）について、商品タイプ固有の分析は**未実装**です。ポジション件数・データ有無の管理とオプション行の共通計算のみです。元仕様書の要件に合わせた追加が必要です。
6. Greeks のファンド全体集計は、全レッグの greeks_basis が確認済みの場合の Weight-adjusted Delta 合計だけです。Notional を使ったリスク量換算は実装していません（定義が不明なため）。
7. 公式の資金フローデータを取り込む欄はありません（推定フローのみ）。
8. NAV 系列の最大 DD は O(n²) の MAXIFS で計算しています。NAV が数万行になると、Excel の再計算に時間がかかる可能性があります。

## 4. データ不足と追加で必要な情報

- 実物の `end_investor_holdings_view.xlsx`（列名・Amount の単位・通貨・as_of 日付・重複行の有無）
- 実物の STEP Composition / Portfolio ファイル（列の構成、Greeks の単位と定義、Weight の定義、Spot が含まれるか、leg を識別する ID があるか）
- Live Fund List の列構成と、fund_status の値
- 元仕様書（各シートの項目、アラート閾値、表示要件、15_QA_Output の定義）
- 各ファンドの `nav_with_distribution` の定義（累積 TR 指数／分配金調整後 NAV／単純な NAV＋分配金）
- fund_exposure_multiplier の根拠資料
- 参照指数（ベンチマーク）として認められる系列と、通貨・ヘッジ条件
- 営業日カレンダー（休日）、営業担当者と宛先の一覧

## 5. テスト結果

LibreOffice で再計算したうえで、Python で独立に計算した期待値と照合しました（`verify.py`。詳細は `docs/test_results.md`）。入力は `fixtures.py` の**架空データ**です。

| # | テスト | 結果 | 主な確認内容 |
|---|---|---|---|
| 1 | fund_code が元データと一致 | PASS | 01 と 07 の順序・値、重複 2 行を保持して検出 |
| 2 | Amount 集約 | PASS | 2 行合算、非数値は空欄＋Check Required、明細を全行保持 |
| 3 | 同一 Index の複数ファンド | PASS | 3 ファンドが同じ Index でもキー・NAV・AUM・リターンは別。曖昧なキーは自動選択しない |
| 4 | Position Raw の非改変 | PASS | 06/04/02 の元の値が入力と一致し、元データ列に数式なし |
| 5 | Ticker 解析 | PASS | P/C・Strike・満期・原資産を正規表現の結果と照合。解析できない Ticker は失敗として表示 |
| 6 | Strike Distance・ITM | PASS | Put/Call 別の判定、Spot の出所、原資産不一致時の STOP |
| 7 | Spot 欠損 | PASS | Spot Missing で距離を計算せず、ゲートは DATA REQUIRED |
| 8 | NAV 欠損 | PASS | リターンは空欄、コメントは【データ不足】。1W・YTD・DD を独立計算と照合 |
| 9 | ポジション差分 | PASS | New/Existing/Matured/Removed/Check Required、Weight 差分、旧版は不使用、Potential Roll |
| 10 | Update Coverage への欠損の反映 | PASS | Missing/Partial/Complete。欠損時に GREEN・READY にならない |
| 11 | 閾値変更の反映 | PASS | 3 種の閾値を変更すると重要度が変わる |
| 12 | Dashboard と Comment の一致 | PASS | コメントは dsh_* の名前付きセルだけを参照し、テキストの数値も一致 |
| 13 | 99_Check の原因表示 | PASS | 未充足の全項目に理由・対象・対応を表示 |
| 14 | 未承認文章の配信不可 | PASS | Reviewed は REVIEW REQUIRED、final_text は空。宛先なしは DATA REQUIRED |
| 15 | 行追加後の数式 | PASS | 追加行が集約と投資家数に反映。評価日より後の NAV は無視。全計算列に calculatedColumnFormula を設定 |
| 16 | 再オープン後の参照維持 | PASS | 保存して再度開いた後も名前定義・テーブル・数式を維持。エラーセルなし（グラフ用の意図的な #N/A を除く） |

**注記**：テストはすべて LibreOffice での結果です。Excel 365 で開いて同じシナリオを確認する受け入れテストを、本番利用の前に実施してください。
