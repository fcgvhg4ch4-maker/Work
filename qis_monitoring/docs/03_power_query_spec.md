# Power Query 取り込み仕様

M コードは `powerquery/` に同梱（ブックには埋め込まれていない：openpyxl では Power Query を生成できないため）。
**Requires Local Excel Configuration**。

## 1. 構成

| クエリ | 種別 | 読込先 | 役割 |
|---|---|---|---|
| `fnReadPositionFile` | 関数 | 接続のみ | 1 ファイルを読み、`tblColMap` で列名・型を標準化。必須列欠落はエラー |
| `qPositionLatest` | 表 | `05_Position_Raw!A5` | フォルダ内で最新の Position Date を持つファイルを取り込む |
| `qPositionImportLog` | 表 | 任意のシート | 全ファイルの診断（エラー・基準日複数・重複行・Ticker 欠損・日付変換不能） |
| `qPositionAll` | 表 | 任意（初期移行用） | アーカイブフォルダ全ファイルの結合（Snapshot ID で重複除去） |
| `qNavOfficial` | 表 | ステージング | 公式 NAV CSV の結合（確認後に 08 へ値で追記） |
| `qMarketCsv` | 表 | ステージング | Bloomberg 未接続時の市場データ CSV（07 の K〜N 列へ貼付け） |

パラメータはブック内のテーブルから読む（パスをコードに埋め込まない）。

- `tblPQParams`（00_Control）：`PositionFolder`、`PositionFilePattern`、`PositionSheetName`、`ArchiveFolder`、`NavFolder`、`MarketCsvPath`
- `tblColMap`（19_Settings E 表）：`SourceColumn`（受領ファイルの列名）→ `StandardColumn`（05 の列名）、`DataType`（date/number/text）、`Required`（Y/N）

## 2. 作成手順（初回のみ）

1. データ ＞ データの取得 ＞ その他のデータソース ＞ 空のクエリ。
2. 詳細エディターに `fnReadPositionFile.pq` を貼り付け、クエリ名を `fnReadPositionFile` にする。
3. 同様に `qPositionLatest` 等を作成。
4. `qPositionLatest` を「読み込み先 ＞ テーブル ＞ 既存のワークシート `05_Position_Raw!$A$5`」。初回は A6:S505 のサンプル値を削除してから読み込む（T 列の取込チェック数式は残す）。
5. `qPositionImportLog` は別シートに読み込み、毎週確認する。
6. クエリのプロパティで「ファイルを開くときにデータを更新する」はオフ（00_Control の手順 3 で明示的に更新する）。

## 3. 取り込みロジック

| 要件 | 実装 |
|---|---|
| 指定フォルダから最新ファイルを取得 | `positions_` で始まる CSV/xlsx のうち、ファイル内 **Position Date の最大値**が最も新しいもの（同日は更新日時が新しいもの）。更新日時だけに頼らない |
| 複数ファンドの分離 | ファイル内は Index Name で識別。ファンドへの展開はブック側の Fund-Index Mapping |
| 列名マッピング | `tblColMap`。運用会社の列名変更時は SourceColumn のみ修正 |
| データ型の統一 | date（yyyy-MM-dd / yyyy/MM/dd / Excel シリアル）、number（en-US）、text（Trim） |
| 重複行検出 | `qPositionImportLog`（基準日・Index・Ticker・open_date・Notional が同一）＋ブック側 `05` T 列・`06` 重複数・DQ11 |
| 基準日の検証 | ファイル内の Position Date が単一か（ログ）。評価基準日との差はブック側（DQ07・DQ08） |
| 取込エラーの記録 | `qPositionImportLog`（try … otherwise でエラー内容を行として残す） |
| 過去ファイルの履歴保存 | `QIS_ArchivePositions`（VBA）で 05 → 06 へ**値で**追記。元ファイルが後で更新されても履歴は変わらない |
| 再取込で履歴を重複させない | Snapshot ID（基準日｜Index Name｜ファイル名）が既存ならスキップ。同一基準日・同一指数で別ファイルが既にある場合も追記せず警告（DQ15 でも検出） |

## 4. ファイル運用

- 受領フォルダ（`PositionFolder`）：最新ファイルを保存。
- アーカイブ（`ArchiveFolder`）：取込後に移動（任意）。`qPositionAll` で全件再構成が可能。
- 命名規則：`positions_YYYYMMDD.csv`（推奨）。基準日はファイル名ではなくファイル内の Position Date で判定する。
