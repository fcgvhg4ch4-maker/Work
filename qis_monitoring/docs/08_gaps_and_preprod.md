# 実装ステータス・未実装機能・本番導入前の確認事項

実装できていない機能を実装済みと表現しないため、区分を明示する。ブック内では `21_Impl_Status` に同じ内容を収録。

## 実装ステータス

| 区分 | 手段 | 内容 |
|---|---|---|
| **Fully Implemented** | Excel 数式 | ポジション標準化（Index ID・Put/Call 判定・Strike Distance・Moneyness・DTE・Position Status）／前回比較（New・Closed・Expired・Increased・Reduced・Possible Roll・Strike Change・ITM 変化）／指数単位集計／NAV 照合 5 区分＋比較不能／TR リターン（週次・月次・YTD・設定来・取得来）／MDD・現在 DD・実現 Vol／Vol Target・Trend/Carry 分析／ルールベースのアラート（29 ルール、閾値は Settings）／営業優先度（スコア内訳表示・手動 Override）／定型コメント（Fund 7 項目・Option 指数 10 項目）／営業担当者別メール本文／データ品質 24 項目／ダッシュボード（絞込・チャート） |
| **Requires Bloomberg Connection** | BDH／BDP | 07 の BDH 列（評価基準日・前週・前月末）・BDP ライブ値・オプション照合（PX_MID・IVOL_MID・OPT_DELTA_MID）・接続テスト。未接続時は NO_CONN となり、MANUAL モードで手入力値を使用。フィールド名・Ticker・単位は端末で要確認。 |
| **Requires Local Excel Configuration** | Power Query | qPositionLatest／qPositionAll／qNavOfficial／qMarketCsv の M コードは powerquery/ に同梱。クエリの作成と 05 への読込設定はローカルの Excel で実施（openpyxl では Power Query を埋め込めないため、ブックには含まれていない）。 |
| **Requires VBA / Outlook Permission** | VBA | vba/modQIS.bas（更新一括実行・ポジション履歴追記（重複防止）・メール履歴保存・Outlook 下書き作成・本文コピー）。.xlsm への保存とマクロ有効化、Outlook オブジェクトモデルへのアクセス許可が必要。 |
| **Requires Additional Data** | データ | Vol Target 項目（Target/Realized Vol・Exposure・上下限・リバランス）と Trend/Carry 項目（Signal・Weight・Carry・Roll Yield・寄与）はポジションファイルにないため、運用会社の週次データが必要（11・12 の入力表）。正式な寄与度データ（09 C 表）。公式 NAV・分配金。 |
| **Future Enhancement** | — | 高度な自然言語生成（社内承認済み AI 環境が前提）／アセットクラス別寄与変化アラート／Greeks の公式ポートフォリオ感応度（単位確認後）／BDS による分配金自動取得／複数シェアクラスの自動照合／ポジション明細の Trade ID（取引識別子）が提供された場合の取引単位比較。 |

## 未実装機能・制約（Future Enhancement を含む）

1. **Power Query はブックに埋め込まれていない。** openpyxl では Power Query（DataMashup）を生成できないため、M コードを `powerquery/` に同梱。ローカル Excel でクエリを作成する必要がある。
2. **VBA はブックに埋め込まれていない。** `vba/modQIS.bas` をインポートして .xlsm で保存する。
3. **Bloomberg 連携は未検証。** 開発環境に Bloomberg 端末がないため、BDH/BDP 数式は `NO_CONN` を返す状態でのみ確認。フィールド名・Ticker・単位・Fill オプションは端末での確認が必要。
4. **Excel 365 での再計算は未検証。** 検証は LibreOffice で実施。Excel 固有の差異（動的配列化による暗黙の交差 `@` の付与など）がないか、初回オープン時に `22_Test_Results` の値と照合する。配列が必要な数式はすべて従来型の配列数式（CSE）で保存している。
5. **実際のポジションファイルは未受領。** 構造分析・テストは追加仕様 A の 17 列定義に基づく架空データで実施（docs/01）。
6. **Greeks の合計は参考値。** 単位・符号の確認前のため、Σ(Greek×Notional) を『参考値（単位未確認）』として表示し、正式なポートフォリオ感応度としては扱わない。
7. **寄与度分析は運用会社データ依存。** 正式データがない場合は Observed（観測）表示に留め、要因を断定しない。Estimated は Vol Target の原資産寄与（前週 Exposure×原指数騰落率）と為替ヘッジなしファンドの為替寄与のみ。
8. **ロール判定は推定。** 同一指数・同一原資産・同一 Put/Call の解消（または満期）と、より遠い満期の新規が同時に起きた場合に Possible Roll とする。取引識別子が提供されれば確定判定に置き換え可能。
9. **Leverage Change／Exposure Change（オプション）**：オプション指数については Gross Notional の前回比として 10 に表示するが、専用アラートは未設定（必要に応じて 19 にルール追加）。
10. **アセットクラス別の損益寄与変化アラート**（Trend/Carry）：寄与データの入力欄はあるが、アラートは未実装（Future Enhancement）。
11. **高度な自然言語生成**：Excel 関数による定型コメントのみ。社内承認済み AI 環境が整備された後に検討（顧客情報を未承認サービスへ送らない）。
12. **容量**：ファンド 30／指数 30／Mapping 30／保有 100／ポジション履歴 1,500 行／NAV 2,000 行／アラート一覧 200 件／メール 1 通あたり 10 保有先。超える場合は build_workbook.py の定数を変更して再生成。

## 本番導入前の確認事項

| # | 区分 | 確認事項 |
|---|---|---|
| 1 | ポジションファイル | 実ファイルのシート名・列名・データ型が仕様（17 列）と一致するか。列名が異なる場合は 19_Settings の tblColMap を修正。 |
| 2 | ポジションファイル | Notional(USD) の符号（買い/売り）と単位（USD 額面か、指数 1 単位あたりか）。 |
| 3 | ポジションファイル | Weight の定義（符号付きか、NAV 比か、指数構成比か）と合計の意味。合計 100% は前提にしていない。 |
| 4 | ポジションファイル | Delta/Gamma/Vega/Theta が 1 単位あたりか、ポジション全体の感応度か。確認後に 02 の『Greeks単位確認済』を Y に変更。 |
| 5 | ポジションファイル | Implied Vol の単位（0.25 か 25 か）。現在は小数（0.25＝25%）を前提。 |
| 6 | ポジションファイル | 『NDX Spot』列が NDX 以外を原資産とする指数でも NDX の値なのか、各原資産の値なのか。確認後に 02 の『NDX Spot列 有効』を設定。 |
| 7 | ポジションファイル | 同一 Ticker の複数行が別取引（open_date 別）か、重複送付か。取引識別子（Trade ID）の提供可否。 |
| 8 | ポジションファイル | Index Level がどの時点（Position Date 終値か、算出日か）の値か。 |
| 9 | ポジションファイル | 満期当日のポジションがファイルに残るか（現在は満期日＝ポジション基準日は有効、満期日＜評価基準日は無効として扱う）。 |
| 10 | Bloomberg | 各 Ticker・フィールド（PX_LAST／PX_MID／IVOL_MID／OPT_DELTA_MID／VOLATILITY_30D／30DAY_IMPVOL_100.0%MNY_DF 等）の存在と単位を端末の FLDS で確認。 |
| 11 | Bloomberg | QIS 指数の Bloomberg Ticker（現在は架空 Ticker）を運用会社・指数算出者に確認。 |
| 12 | Bloomberg | BDH の Fill オプション（評価基準日は Fill=N、前週・前月末は Fill=P）の挙動を休日で確認。 |
| 13 | Bloomberg | openpyxl で書いた BDH/BDP 数式が Add-in 読込済みの Excel で認識されるか（認識されない場合はセルを再入力または Bloomberg の Refresh）。 |
| 14 | Bloomberg | データ保存・社内共有・メール記載の可否を Bloomberg 契約（データライセンス）で確認。顧客向け転送は行わない運用。 |
| 15 | NAV | 公式 NAV の受領形式・公表ラグ（02 の NAV公表ラグ）・Share Class・NAV Type・Basis（TR/Price）。 |
| 16 | NAV | 分配金の Ex-Date と支払日、TR 計算での扱い（現在は Ex-Date で NAV に加算）。 |
| 17 | NAV | NAV 照合の許容差（19_Settings：MATCH 0.05%・WARNING 0.3%）が妥当か。 |
| 18 | 閾値 | 全アラート閾値（19_Settings A 表）は仮の初期値。商品特性に合わせて調整し、変更履歴を残す。 |
| 19 | 閾値 | 営業優先度のスコア配点と High/Medium の境界。 |
| 20 | 休日 | 19_Settings F 表の休日（米国市場）を最新化。日本の祝日を使う判定が必要か。 |
| 21 | Power Query | フォルダパス（tblPQParams）、ファイル命名規則、CSV の文字コード（UTF-8 BOM 前提）、Excel ファイルのシート名。 |
| 22 | VBA / Outlook | マクロの配布・署名ポリシー、Outlook オブジェクトモデルのアクセス許可（下書き作成のみ。送信しない）。 |
| 23 | 顧客情報 | 地銀・保有残高等の顧客情報の保管場所とアクセス権。外部の未承認サービスへ送信しない。 |
| 24 | 監査 | 17_Email_History・06_Position_History・08_NAV_History の保存期間と、ブックのバージョン管理（S_WorkbookVersion）。 |
| 25 | 性能 | ポジション履歴 1,500 行・NAV 履歴 2,000 行の容量で足りるか。超える場合は容量拡張か、古い履歴の別ブック退避。 |
| 26 | 運用テスト | 実データで 2〜4 週間の並行運用を行い、手計算・運用会社レポートとの突合を実施。 |
