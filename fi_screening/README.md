# 全国金融機関ターゲット分析データベース（米国債ラダー × Dynamic FX Hedging）

公開情報から、地方銀行・第二地方銀行・信用金庫・信金中央金庫（計350先）の有価証券運用・外債関連データを
出典付きで収集し、スコアリングと営業仮説の検証に使う Excel ブックを再現可能に生成する。

- 設計書（承認済み）: [`DESIGN.md`](DESIGN.md)
- 成果物: `dist/FI_Target_Screening.xlsx`（12シート＋M1受け皿）
- 正本データ: `data/*.csv`（ブックは毎回ここから再生成する）

## パイプライン

取得物（PDF・HTML・Excel）はリポジトリに入れず、任意の作業ディレクトリ `<raw>` に保存する。

```sh
RAW=/path/to/raw
python3 -I fi_screening/scripts/fetch_sources.py   $RAW   # 協会名簿・全信協一覧・全銀協 行別財務諸表(3期)
python3 -I fi_screening/scripts/build_data.py      $RAW   # universe.csv / financial_data.csv / sources.csv / batch_log.csv
python3 -I fi_screening/scripts/crawl_disclosure.py $RAW <業態|all> [code...]   # 各機関サイト→ディスクロ誌PDF取得
python3 -I fi_screening/scripts/extract_securities.py $RAW   # 銀行: 外国債券等を抽出（全銀協計数と突合）
python3 -I fi_screening/scripts/extract_shinkin.py   $RAW   # 信金: 種類別残高・預金等を抽出（内部整合で検証）
python3 -I fi_screening/scripts/ingest_extracted.py         # 抽出値→auto_financial.csv、一致した全銀協計数をVERIFIEDへ
python3 -I fi_screening/scripts/extract_policy.py    $RAW   # 運用方針の原文候補（policy_candidates.csv）
python3 -I fi_screening/scripts/export_xlsx.py              # Excel 生成
```

## データファイル

| ファイル | 内容 | 作成 |
| --- | --- | --- |
| `universe.csv` | 母集団350先（主キー: 金融機関コード） | build_data |
| `sources.csv` / `auto_sources.csv` / `manual_sources.csv` | 出典台帳 | 自動／手入力 |
| `financial_data.csv` | 全銀協 行別単体計数（地銀・第二地銀、FY2023-25） | build_data |
| `auto_financial.csv` | ディスクロ誌からの機械抽出値（検証条件を満たしたもののみ） | ingest_extracted |
| `manual_financial.csv` | 分析者の手入力値（最優先で採用） | 手入力 |
| `extracted_foreign.csv` / `extracted_shinkin.csv` | 抽出の中間結果（表の所在・検証内容） | extract_* |
| `disclosure_seeds.csv` | クローラの起点URL（持株会社サイト等） | 手入力 |
| `policy_candidates.csv` | 運用方針の原文候補（未精査） | extract_policy |
| `policy_evidence.csv` | 採用した運用方針の証拠（精査済み） | 手入力 |
| `qual_inputs.csv` | 定性評価（A3/A4/B1-B3/D1-D3 の点数と根拠ID） | 手入力 |
| `selection.csv` / `profiles.csv` | 選定段階・理由、銀行別プロファイル | 手入力 |

## 検証ルール（誤読防止）

- **銀行**：ディスクロ誌の期末残高表の「合計」が全銀協の有価証券（同期・単体）と一致し、かつ国債・地方債・株式の
  2項目以上が一致した列だけを採用。億円表示は百万円換算（±100百万円の丸め誤差を許容し、その旨を記録）。
- **信金**：残存期間別表の種類別合計が、経営指標の有価証券残高と一致した年度だけを採用（単位は千円/百万円を自動判定）。
- 一致しない表、外国証券の内訳行がない表からは値を採らず `NOT YET CHECKED` とする（非開示と区別）。
- 「外国証券」は外国債券＋外国株式（表に内訳がある場合）。開示が無い場合の合算値は `ESTIMATED`（式を記録）。
