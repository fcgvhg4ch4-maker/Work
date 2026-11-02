// qMarketCsv: Bloomberg 未接続時の代替入力（CSV: Key, Ticker, Field, Current, Prev Week, Prev Month End, As-of）。
// 読込先: ステージング用シート。07_Bloomberg_Data の K〜N 列（手入力欄）へ値で貼付け。
// As-of が評価基準日と異なる値は、07 で自動的に STALE となり最新値として扱われない。
let
    Params = Excel.CurrentWorkbook(){[Name = "tblPQParams"]}[Content],
    P      = (name as text) => Params{[Param = name]}[Value],
    Src    = Csv.Document(File.Contents(P("MarketCsvPath")), [Delimiter = ",", Encoding = 65001]),
    Hdr    = Table.PromoteHeaders(Src, [PromoteAllScalars = true]),
    Typed  = Table.TransformColumnTypes(Hdr, {{"Current", type number}, {"Prev Week", type number}, {"Prev Month End", type number}, {"As-of", type date}}, "en-US")
in
    Typed
