# テスト結果（自動生成：verify.py）

- 実行日時: 2026-11-02 02:59
- 再計算エンジン: LibreOffice（headless）。Excel 365 本体での再計算は未実施（下記注記参照）。
- 入力: fixtures.py の架空テストデータ（TEST）

| # | テスト | 結果 | 詳細 |
|---|---|---|---|
| 1 | Live Fund の fund_code が元データと一致 | PASS | 3/3 assertions |
| 2 | Investor Holdings の Amount 集約 | PASS | 9/9 assertions |
| 3 | 同じ Index の複数ファンドを区別 | PASS | 7/7 assertions |
| 4 | Position Raw の元データが改変されない | PASS | 4/4 assertions |
| 5 | BloombergTicker の Put/Call・Strike 解析 | PASS | 21/21 assertions |
| 6 | STEP の Strike Distance と ITM 判定 | PASS | 78/78 assertions |
| 7 | Spot 欠損時に Strike Distance を計算しない | PASS | 4/4 assertions |
| 8 | NAV 欠損時に架空のリターンを表示しない（＋リターン計算の独立検証） | PASS | 14/14 assertions |
| 9 | ポジション差分の計算 | PASS | 11/11 assertions |
| 10 | データ欠損が Update Coverage に反映 | PASS | 6/6 assertions |
| 11 | アラート閾値の変更が反映 | PASS | 3/3 assertions |
| 12 | Dashboard と Comment Output の数値一致 | PASS | 7/7 assertions |
| 13 | 99_Check でエラー原因を確認できる | PASS | 5/5 assertions |
| 14 | 未承認の文章が配信可能と判定されない | PASS | 5/5 assertions |
| 15 | テーブルに行を追加しても数式が機能 | PASS | 4/4 assertions |
| 16 | 再オープン後も数式参照が維持 | PASS | 5/5 assertions |
