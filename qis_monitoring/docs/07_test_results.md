# テスト結果（架空データ）

- 実行日時: 2026-11-02 04:24
- 合計 83 件 / 合格 83 件 / 不合格 0 件
- 方法: `build/verify_workbook.py` がブックを LibreOffice（ヘッドレス）で再計算し、`sample_data.py` から Python で独立に計算した期待値と照合。派生シナリオ（V*）はブックの入力セルを書き換えて再計算。
- 注意: LibreOffice での再計算結果であり、Excel 365 での再計算は本番前に別途確認すること（docs/08）。VBA・Power Query・Bloomberg 接続はこの検証の対象外。

| ID | シナリオ | 確認内容 | 期待値 | 結果値 | 判定 |
|---|---|---|---|---|---|
| T01-1 | NAV下落 | F001 週次リターン（TR） | -0.0279993 | -0.0279993 | PASS |
| T01-2 | NAV下落 | F001 月次リターン（TR） | -0.053029 | -0.053029 | PASS |
| T01-3F002 | NAV下落 | F002 週次リターン（分配金調整後 TR） | -0.0159346 | -0.0159346 | PASS |
| T01-3F003 | NAV下落 | F003 週次リターン（分配金調整後 TR） | -0.012 | -0.012 | PASS |
| T01-3F004 | NAV下落 | F004 週次リターン（分配金調整後 TR） | -0.00202834 | -0.00202834 | PASS |
| T01-3F005 | NAV下落 | F005 週次リターン（分配金調整後 TR） | -0.0100402 | -0.0100402 | PASS |
| T01-4 | NAV下落 | F004 現在ドローダウン（TR） | -0.139581 | -0.139581 | PASS |
| T01-5 | NAV下落 | F001 実現Vol（26週・年率） | 0.0859702 | 0.0859702 | PASS |
| T02-F001 | NAV不一致 | F001 NAV照合結果 | MATCH | MATCH | PASS |
| T02-F002 | NAV不一致 | F002 NAV照合結果 | MISMATCH | MISMATCH | PASS |
| T02-F003 | NAV不一致 | F003 NAV照合結果 | WARNING | WARNING | PASS |
| T02-F004 | NAV不一致 | F004 NAV照合結果 | DATE DIFFERENCE | DATE DIFFERENCE | PASS |
| T02-F005 | NAV不一致 | F005 NAV照合結果 | MISSING DATA | MISSING DATA | PASS |
| T01-6 | NAV下落 | F001 週次NAV下落アラート＝High | High | High | PASS |
| T01-7 | NAV下落 | F001 月次NAV下落アラート＝High | High | High | PASS |
| T01-8 | NAV下落 | F004 ドローダウン拡大アラート＝Medium | Medium | Medium | PASS |
| T02-6 | NAV不一致 | F002 NAV不一致アラート＝High | High | High | PASS |
| T02-7 | NAV不一致 | F004 日付ずれは差額アラートを出さず『照合不能』 | ('Medium', None) | ('Medium', None) | PASS |
| T03-1 | 新規ポジション | NDX 01/15/27 P19000 → New | New | New | PASS |
| T03-2 | 新規ポジション | NDX 11/04/26 C21000 → New（CALL 判定） | ('New', 'CALL') | ('New', 'CALL') | PASS |
| T03-3 | 新規ポジション | 全銘柄の分類が独立計算と一致 | {'NDX US 11/06/26 P19600 Index': 'Closed', 'NDX US 12/18/26 P18500 Index': 'Unchanged', 'NDX US 11/20/26 P19500 Index': 'Increased', 'NDX US 11/04/26 C21000 Index': 'New', 'NDX US 12/18/26 P20200 Index': 'Unchanged', 'NDX US 10/30/26 P19800 Index': 'Expired', 'NDX US 01/15/27 P19000 Index': 'New'} | {'NDX US 10/30/26 P19800 Index': 'Expired', 'NDX US 11/06/26 P19600 Index': 'Closed', 'NDX US 11/20/26 P19500 Index': 'Increased', 'NDX US 12/18/26 P20200 Index': 'Unchanged', 'NDX US 12/18/26 P18500 Index': 'Unchanged', 'NDX US 01/15/27 P19000 Index': 'New', 'NDX US 11/04/26 C21000 Index': 'New'} | PASS |
| T04-1 | ポジション解消 | NDX 11/06/26 P19600 → Closed | Closed | Closed | PASS |
| T04-2 | ポジション解消 | NDX 10/30/26 P19800 → Expired（満期到来） | Expired | Expired | PASS |
| T04-3 | ポジション解消 | 解消＋新満期の新規 → Possible Roll（推定） | ('Possible Roll (旧側)', 'Possible Roll (新規側)') | ('Possible Roll (旧側)', 'Possible Roll (新規側)') | PASS |
| T04-4 | ポジション解消 | Call の新規はロール候補にしない（Put/Call 別グループ） | None | None | PASS |
| T04-5 | ポジション解消 | 同一銘柄2取引（open_date 別）の集約：P18500 前回/今回 Notional | (40000000, 40000000) | (40000000, 40000000) | PASS |
| T05-1 | Strike接近 | IDX_NDXPW Nearest Strike Distance | 0.0074813 | 0.0074813 | PASS |
| T05-2 | Strike接近 | OTM→ITM 変化銘柄 | ['NDX US 12/18/26 P20200 Index'] | ['NDX US 12/18/26 P20200 Index'] | PASS |
| T05-3 | Strike接近 | Strike接近（3%以内）＝High、5%ルールは抑制 | ('High', None) | ('High', None) | PASS |
| T05-4 | Strike接近 | ITMへの変化アラート＝High | High | High | PASS |
| T05-5 | Strike接近 | IDX_NDXPS（最接近 5.24%）は Strike アラートなし | False | False | PASS |
| T05-6 | Strike接近 | 最短満期までの営業日数（休日考慮） | 3 | 3 | PASS |
| T05-7 | Strike接近 | 満期接近（3営業日以内）＝Medium、30日ルールは抑制 | ('Medium', None) | ('Medium', None) | PASS |
| T05-8 | Strike接近 | 加重平均IV変化（\|Notional\|加重） | 0.0443636 | 0.0443636 | PASS |
| T05-9 | Strike接近 | Gamma×Notional 変化率（参考値） | 0.508532 | 0.508532 | PASS |
| T05-10 | Strike接近 | 有効ポジション数・Gross Notional | (7, 165000000) | (7, 165000000) | PASS |
| T05-11 | Strike接近 | Greeks は単位確認前『参考値』表示 | 参考値（単位未確認） | 参考値（単位未確認） | PASS |
| T05-12 | Strike接近 | 前回ファイルの満期到来ポジション（10/30 満期 P19800）は今回の有効明細に含まれない | False | False | PASS |
| T06-1 | Vol Target Exposure縮小 | Exposure 変化率 | -0.29661 | -0.29661 | PASS |
| T06-2 | Vol Target Exposure縮小 | Exposure急縮小＝High／実現Vol急上昇＝Medium | ('High', 'Medium') | ('High', 'Medium') | PASS |
| T06-3 | Vol Target Exposure縮小 | コメントに『観測』としての説明 | True | True | PASS |
| T07-1 | Trendシグナル反転 | シグナル反転件数／L/S反転件数 | (1, 1) | (1, 1) | PASS |
| T07-2 | Trendシグナル反転 | 反転銘柄の特定 | Equity:S&P500 E-mini(+0.60→-0.20) | Equity:S&P500 E-mini(+0.60→-0.20) | PASS |
| T07-3 | Trendシグナル反転 | 加重Carry変化 | -0.0108049 | -0.0108049 | PASS |
| T07-4 | Trendシグナル反転 | Trendシグナル反転＝Medium・Carry低下＝Low | ('Medium', 'Low') | ('Medium', 'Low') | PASS |
| T08-1 | Bloombergデータ欠損 | SX5E 欠損 → MISSING | MISSING | MISSING | PASS |
| T08-2 | Bloombergデータ欠損 | JGB10 As-of 10/28 → STALE（値は採用しない） | ('STALE', None) | ('STALE', None) | PASS |
| T08-3 | Bloombergデータ欠損 | QIS_TC 欠損 → F004 参照指数週次は空欄（前回値で代用しない） | None | None | PASS |
| T08-4 | Bloombergデータ欠損 | 市場データ欠損アラート（IDX_TC） | Medium | Medium | PASS |
| T08-5 | Bloombergデータ欠損 | Add-in 未接続時の BDH は NO_CONN | NO_CONN | NO_CONN | PASS |
| T08-6 | Bloombergデータ欠損 | コメントに欠損の明示 | True | True | PASS |
| T08-7 | Bloombergデータ欠損 | DQ16（重要市場データ欠損）＝OPEN | ('NG', 'OPEN') | ('NG', 'OPEN') | PASS |
| T09-1 | 複数地銀・複数ファンド | 優先度一覧の行数（保有 7 件） | 7 | 7 | PASS |
| T09-2 | 複数地銀・複数ファンド | 先頭は B銀行（NAV不一致・償還接近・接触超過） | ('B銀行', 'High') | ('B銀行', 'High') | PASS |
| T09-3 | 複数地銀・複数ファンド | Priority の内訳（High/Medium） | {'High': 4, 'Medium': 3} | {'High': 4, 'Medium': 3} | PASS |
| T09-4 | 複数地銀・複数ファンド | アラートの Investor に保有地銀を列挙 | A銀行、D銀行 | A銀行、D銀行 | PASS |
| T09-5 | 複数地銀・複数ファンド | Fund-Index 1対多（F002 は 2 指数のアラートを持つ） | True | True | PASS |
| T09-6 | 複数地銀・複数ファンド | 保有先なしファンド（F005）を DQ20 で検出 | ('NG', 'F005') | ('NG', 'F005') | PASS |
| T10-1 | メール生成 | 件名フォーマット | 【QIS Weekly Update】担当先ファンド状況｜2026/11/02 | 【QIS Weekly Update】担当先ファンド状況｜2026/11/02 | PASS |
| T10-2 | メール生成 | 宛名・担当先（A銀行・B銀行）のみ記載 | (True, True, True, False, False) | (True, True, True, False, False) | PASS |
| T10-3 | メール生成 | 重大DQ未解決 → 本文冒頭に【要確認】・ステータス要確認 | (True, True) | (True, True) | PASS |
| T10-4 | メール生成 | アラートのある先が上位（B銀行→A銀行の順） | True | True | PASS |
| T10-5 | メール生成 | 必須セクションの存在 | True | True | PASS |
| T10-6 | メール生成 | 社内向け質問（（社内）〜）は顧客質問欄に出さない | False | False | PASS |
| T10-7 | メール生成 | 顧客転送不可の注記 | True | True | PASS |
| T10-8 | メール生成 | 本文の長さ（1分で読める目安：3,000字以内） | True | True | PASS |
| H-1 | 数式健全性 | 全シートにエラー値（#VALUE! 等）がない（チャート補助の #N/A を除く） | 0 | 0 | PASS |
| V1-1 | 同一ファイル再取込 | 重複行を DQ11 で検出（Critical） | ('NG', 18) | ('NG', 18) | PASS |
| V1-2 | 同一ファイル再取込 | メールは『要確認』 | 要確認 | 要確認 | PASS |
| V2-1 | Bloomberg 未接続で BLOOMBERG モード | 全市場データが MISSING（手入力値・古い値を使わない） | True | True | PASS |
| V2-2 | Bloomberg 未接続で BLOOMBERG モード | 参照指数週次は全ファンド空欄 | True | True | PASS |
| V2-3 | Bloomberg 未接続で BLOOMBERG モード | NAV（公式）由来のリターンは影響なし | True | True | PASS |
| V3-1 | 公式NAV欠損 | 週次リターンは空欄（前週 NAV で代用しない） | None | None | PASS |
| V3-2 | 公式NAV欠損 | DQ01 で検出 | ('NG', 'F001') | ('NG', 'F001') | PASS |
| V3-3 | 公式NAV欠損 | コメントは【データ欠損】を明示 | True | True | PASS |
| V3-4 | 公式NAV欠損 | 公表予定日（WORKDAY+1）超過で NAV公表遅延アラート | True | True | PASS |
| V4-1 | DQ 承認 | Critical を Ack（確認者記入）→ メール『レビュー可』 | ('OK', 'レビュー可') | ('OK', 'レビュー可') | PASS |
| V4-2 | DQ 承認 | 本文の【要確認】行が消える | False | False | PASS |
| V5-1 | 担当者切替 | 佐藤 花子 → C銀行・E銀行のみ | (True, True, True, False, False) | (True, True, True, False, False) | PASS |
| V5-2 | 担当者切替 | 宛先メールアドレス | hanako.sato@example.com | hanako.sato@example.com | PASS |
| V6-1 | 手動変更 | Override 列で優先度を変更・Reason に（手動変更） | ('Low', True) | ('Low', True) | PASS |
| V6-2 | 手動変更 | Override したコメントが Final に反映 | （テスト）手動修正したリスクコメント | （テスト）手動修正したリスクコメント | PASS |
| V7-1 | 閾値変更 | 週次下落閾値を -3% に変更 → F001（-2.80%）のアラートが消える | False | False | PASS |
