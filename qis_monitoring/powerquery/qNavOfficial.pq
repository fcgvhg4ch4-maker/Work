// qNavOfficial: 運用会社の公式 NAV ファイル（CSV: Fund ID, NAV Date, NAV, Currency, Source, NAV Type, Basis, Share Class）を結合。
// 読込先: ステージング用シート。内容確認後、08_NAV_History の A〜H 列へ値で追記する（重複は 08 の「重複」列で検出）。
let
    Params   = Excel.CurrentWorkbook(){[Name = "tblPQParams"]}[Content],
    P        = (name as text) => Params{[Param = name]}[Value],
    Files    = Table.SelectRows(Folder.Files(P("NavFolder")), each Text.Lower([Extension]) = ".csv"),
    Parsed   = Table.AddColumn(Files, "Data", each Table.PromoteHeaders(Csv.Document([Content], [Delimiter = ",", Encoding = 65001]), [PromoteAllScalars = true])),
    Combined = Table.Combine(Parsed[Data]),
    Typed    = Table.TransformColumnTypes(Combined, {{"NAV Date", type date}, {"NAV", type number}}, "en-US"),
    Dedup    = Table.Distinct(Typed, {"Fund ID", "Source", "NAV Date"})
in
    Dedup
