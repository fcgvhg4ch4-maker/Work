# 全国金融機関ターゲット分析（米国債ラダー × Dynamic FX Hedging）

全国の地方銀行・第二地方銀行・信用金庫・信金中央金庫を対象に、公開情報から営業仮説を検証するためのデータベース。

**状態：Phase 0（設計・承認待ち）。個別金融機関のデータはまだ 1 件も収集していない。**

| パス | 内容 |
| --- | --- |
| `docs/01_research_design.md` | 調査設計書（母集団・情報源・データ項目・業態別方法・スコアリング・Excel 構造・バッチ計画） |
| `build_workbook.py` | `data/*.csv` から Excel ブックを生成 |
| `data/*.csv` | 正本データ（公開情報のみ）。現在はヘッダのみ |
| `dist/FI_Target_Screening_template.xlsx` | 空テンプレート（12 シート＋README / Metric Master / Config / Lists） |

```sh
python3 build_workbook.py --template   # 空テンプレート
python3 build_workbook.py              # data/*.csv を読み込んだブック（dist/FI_Target_Screening.xlsx）
```

## 社内情報の扱い

営業カバレッジ・商談履歴・M1 フローなど社内情報は**このリポジトリにコミットしない**。生成したブックの
`10 Sales Coverage Input` / `10b M1 Flow Input` / `11 Sales Action Plan`（赤いタブ）に直接記入する。
`data/internal/` は `.gitignore` 済み。

## 検証

架空のテストデータ（スクラッチ領域のみ、リポジトリには含めない）でブックを生成し、LibreOffice で再計算して
次を確認した：単位換算（億円→百万円）、未調査・非開示・不一致・重複の文字表示、ピア内パーセンタイルと配点、
信金の規模区分、重複コード検知、ランキング（充足率による除外）、品質ダッシュボードの件数、数式エラー 0 件。
