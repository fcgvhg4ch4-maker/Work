# QIS Fund Monitoring & Sales Update Workbook — 操作説明（営業担当者向け）

対象ファイル：`dist/QIS_Fund_Monitoring.xlsx`（本番用テンプレート。データは空）
検証用：`dist/QIS_Fund_Monitoring_TESTDATA.xlsx`（**架空データ**。顧客対応には絶対に使用しないこと）

## 0. 基本ルール

| セルの色 | 意味 | 触ってよいか |
|---|---|---|
| 淡いグレー | 元データ（受領ファイルの値） | 値貼り付けのみ。手で値を書き換えない |
| 淡い黄色 | 入力セル | 入力する |
| 淡い青 | 計算セル | **触らない**（数式） |

- マクロ・VBA は使っていません。自動メール送信もしません。
- 元データは「値として貼り付け」（Ctrl+Alt+V → 値）。書式ごと貼り付けると日付が文字列になることがあります。
- 行を追加するときは**テーブルの最終行の直下**に貼り付けてください。テーブルが自動で広がり、青い列の数式も自動で入ります。
- 古いデータ（NAV・ポジション）は**削除しない**でください。最新版の判定はワークブックが行います。
- 「—」や空欄は「データなし」です。0 ではありません。
- Excel のオプションで「計算方法＝自動」になっていることを確認してください。

## 1. 新しいファンドを追加する

1. `01_Live_Fund_Master` の `tbl_Fund` 最終行の下に、Live Fund List の行をグレー列の順で貼り付ける。
2. `03_Strategy_Map` の `tbl_FundMap` に同じ `fund_code` の行を追加し、次のいずれかを設定する。
   - ① `explicit_strategy_data_key`（fund_code に対して明示的に決まっている場合）
   - ② 右表 `tbl_IndexMap` に Index → Strategy Data Key の承認済み行（approved=Yes）を追加
   - ③ `manual_strategy_data_key` ＋ `mapping_approved_by`
3. `mapping_status` が `Mapped` になったことを確認する。`Review Required` / `Mapping Required` の場合は理由が表示されます。
   同じ Index に承認済みキーが複数ある場合、ワークブックはキーを自動で選びません。① の explicit key を設定してください。
4. 参照指数を使う場合は `benchmark_series_id` を設定する。通貨・ヘッジ条件を確認した場合のみ `benchmark_basis_confirmed=Yes`。
5. `fund_exposure_multiplier` は根拠（`multiplier_basis`）があり `multiplier_confirmed=Yes` の場合だけ計算に使われます。未確認なら空欄のままにしてください（1.0 とはみなしません）。
6. 分配金込みリターンを出す場合は `nav_dist_definition` を運用会社に確認して選択し、`nav_dist_def_confirmed=Yes` を入力する。
7. ファンド数がファンド容量（60）を超えると `99_Check` の C17 が DATA REQUIRED になります → 管理者が `build_workbook.py` の `CAP_FUND` を増やしてワークブックを作り直してください。

## 2. 新しい投資家を追加する

1. `02_End_Investor_Holdings` 左表 `tbl_HoldRaw` に、end_investor_holdings_view の行を貼り付ける（全行。集約は自動）。
2. 右表 `tbl_HoldAgg` に「1投資家×1ファンド＝1行」の集約が自動で作られる。Amount に数値以外の値があると集約額は空欄になり、`Check Required: Amount` と表示されます。
3. `12_Distribution_List` の `tbl_Dist` に `end_investor_bdr_id` を1行追加し、宛先（to/cc）・配信方式（Individual / Combined）・担当営業（必要なら上書き）を入力する。メールアドレスが分からない場合は**空欄のまま**にしてください（配信判定は DATA REQUIRED のまま）。
4. 登録漏れは `99_Check` の C07 に表示されます。

## 3. NAV を更新する

1. `04_NAV_Input` の `tbl_NAV` の最終行の下に、新しい日付の NAV 行を貼り付ける（`fund_code, nav_date, nav, nav_with_distribution, distribution_amount, aum, …`）。
2. 同じ fund_code × nav_date が2行あると、両方とも計算から除外されます（`row_status = Duplicate date`、99_Check C11）。正しい行を元ファイルで確認してから、誤っている行を削除してください。
3. `00_Control` の評価日を更新する。NAV 基準日は「評価日以前で最も新しい有効な NAV 日」です。
4. `07_NAV_Calc` で `nav_status` が `OK` になっていることを確認する（`Stale` は遅延。閾値は Control）。

## 4. 市場データを更新する

1. `05_Market_Input` の `tbl_Mkt` に `series_id`（Bloomberg 形式。例：`NKY Index`）、`data_date`、`value` などを追加する。
2. STEP の Spot は、ティッカーから解析した原資産（例：`NKY Index`）と同じ `series_id` の系列を使います。
   ティッカー上の表記と Spot 系列名が違う場合（例：`SPX US Index` → `SPX Index`）は、右表 `tbl_UndlMap` に対応を登録して `confirmed=Yes` を入力してください。

## 5. STEP ポジションを取り込む

1. `06_Position_Raw` の `tbl_PosRaw` 最終行の下に、Composition（または Portfolio）ファイルの行を貼り付ける。
2. 各行に必ず次の値を入れる：`source_type`（Composition / Portfolio）、`strategy_data_key`、`position_date`、`received_datetime`（受領日時）、`file_name`、`file_seq`（取込連番）。
3. 同じ日のファイルを再受領した場合も、古い行は削除せずに追加する。`received_datetime` が最大（同じ場合は `file_seq` が最大）のファイルが最新版として使われ、古い版は `Superseded (kept)` と表示されます。
4. 同じ日に Composition がある場合は Composition を使用し、Portfolio は照合（`08 port_recon`）に使います。
5. 同じ Ticker・同じ Date of Entry のレッグが複数ある場合は、`leg_id` 列に識別子を入れてください（入れないと `Check Required: non-unique key`）。
6. `08_Position_Calc` の `calc_status` を確認する。
   - `STOP: Spot/Strike underlying mismatch`：Spot の原資産がストライクの原資産と一致しないため計算を停止。
   - `Spot Missing` / `Spot Stale`：05 に Spot を追加。
   - `Parse Failed`：ティッカーを解析できない。原本の put_call_raw / strike_raw / maturity_raw が空なら、元ファイルを確認する。
7. ポジション行がポジション容量（1,500 行）を超えると `99_Check` の C16 が DATA REQUIRED になります → 古い履歴を別ブックに退避するか、容量を増やしたワークブックを作り直してください。

## 6. Dashboard を更新する

1. `00_Control` で評価日（C8）、対象投資家（C9）、対象 fund_code（C10）を選択する。
2. `13_Dashboard` が自動で切り替わる。上段：投資家・ファンド・基準日・ステータス、中段：NAV/パフォーマンス・参照指数比較・ポジション、下段：営業向けサマリー。
3. 参照指数は `Benchmark (confirmed)` の場合のみ差分（Fund − Benchmark）を表示します。それ以外は参考値です。
4. 印刷：A4 横・1 ページ幅に設定済み（ファイル → 印刷 → PDF）。

## 7. コメント案をレビューする

1. `14_Comment_Output` でドラフト（件名〜データに関する注意事項）を確認する。【データ不足】と表示された箇所は、推測で書き足さずに元データを補完してください。
2. `comment_id` をコピーし、下表 `tbl_CommentReview` に新しい行を作って `comment_id` を貼り付ける。
3. `reviewed_text` にドラフトを**値として**貼り付け、必要に応じて修正する。`reviewer` と `review_status=Reviewed` を入力する。
4. 承認者が内容を確認し、`review_status=Approved`、`approver`、`approved_date` を入力する。
5. `final_text` に承認済みの文章が表示されます（承認されていない文章は表示されません）。
6. `comment_id` には評価日・NAV 基準日・ポジション基準日が含まれます。データを更新すると `comment_id` が変わり、**再レビュー・再承認**が必要になります。
7. `15_QA_Output` の想定 Q&A もレビューし、結果を `qa_review_status` に入力する。

## 8. 配信可否を確認する

1. `99_Check` の最終判定（C8）を確認する。
   - **READY**：必須データ・マッピング・レビュー・承認・宛先がすべて揃っている。
   - **REVIEW REQUIRED**：レビュー・承認・重要な確認（High アラートの確認、閾値の照合など）が未完了。
   - **DATA REQUIRED**：必須データの欠損、マッピング不備、計算不能、宛先なし、元データの矛盾がある。
2. OK 以外の行の Error Reason / Affected Fund / Required Action を見て、元シートを修正する。
3. `10_Alert` の High アラートは、`alert_id` を右表 `tbl_AlertReview` にコピーして確認結果を入力する（未確認の High があると REVIEW REQUIRED）。
4. `12_Distribution_List` で投資家単位の配信判定を確認する。Combined（統合配信）は保有する全ファンドが READY の場合のみ READY。Individual（個別配信）は右表 `tbl_DistDetail` で READY の行だけを配信してください。
5. 配信後は `last_sent_date`・`distribution_status=Sent`・`next_scheduled_date` を入力する。
6. 初回利用時：`00_Control` の閾値は暫定値です。元仕様書の値を入力し、C12（閾値照合済み）を `Yes` にするまで READY にはなりません。

## 付録 A. Power Query を使う場合（任意・未実装）

本ワークブックには Power Query の接続は含まれていません（貼り付けで更新する構造）。Power Query を使う場合は次の方針で、各元データテーブルの**グレー列だけ**を置き換えてください。

- 取り込み元のパス：共有フォルダ上の受領ファイル（例：`\\share\QIS\NAV\*.csv`）をフォルダー接続で読み込む。パスは Control に記載して管理する。
- 更新手順：データ → すべて更新。更新後に 99_Check を確認。
- 重複処理：クエリ内で重複を削除しない（重複はワークブック側で検出・表示する）。
- 履歴保持：フォルダー内の過去ファイルを削除しない。`file_name`・`received_datetime` をクエリで付与し、最新版の判定はワークブックの数式に任せる。

## 付録 B. 列名が実ファイルと異なる場合

`end_investor_holdings_view.xlsx` などの実ファイルの列構成は、本ワークブック作成時点で入手できていません。グレー列の名前・順番は仮の設計です。実ファイルの列の意味を確認したうえで、グレー列の順番に並べ替えてから貼り付けてください（意味の異なる列を推測で割り当てないこと）。
