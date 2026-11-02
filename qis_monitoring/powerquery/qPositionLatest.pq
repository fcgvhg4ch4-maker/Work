// qPositionLatest: 指定フォルダの positions_* ファイルのうち、最新の Position Date を含むファイルを読み込む。
// 読込先: 05_Position_Raw!A5（既存の表を置き換える。列順は 05 と同じ）
let
    Params      = Excel.CurrentWorkbook(){[Name = "tblPQParams"]}[Content],
    P           = (name as text) => Params{[Param = name]}[Value],
    Files       = Folder.Files(P("PositionFolder")),
    Candidates  = Table.SelectRows(Files, each Text.StartsWith(Text.Lower([Name]), "positions_")
                                          and List.Contains({".csv", ".xlsx"}, Text.Lower([Extension]))
                                          and not Text.StartsWith([Name], "~$")),
    Loaded      = Table.AddColumn(Candidates, "Data", each fnReadPositionFile([Content], [Extension], P("PositionSheetName"), [Name])),
    WithMaxDate = Table.AddColumn(Loaded, "MaxPosDate", each List.Max([Data][Position Date]), type date),
    // 最新基準日。同じ基準日のファイルが複数ある場合は更新日時が最新のもの（衝突は qPositionImportLog に記録）
    Sorted      = Table.Sort(WithMaxDate, {{"MaxPosDate", Order.Descending}, {"Date modified", Order.Descending}}),
    Latest      = Table.FirstN(Sorted, 1),
    Expanded    = Table.Combine(Latest[Data]),
    Stamped     = Table.AddColumn(Expanded, "Imported At", each DateTime.LocalNow(), type datetime),
    Ordered     = Table.SelectColumns(Stamped,
                    {"Position Date", "Index Name", "Name", "BloombergTicker", "Option Mid Price", "Strike", "Maturity Date",
                     "Notional(USD)", "Delta", "Gamma", "Vega", "Theta", "Implied Vol", "NDX Spot", "Index Level", "Weight",
                     "open_date", "Source File", "Imported At"}, MissingField.UseNull)
in
    Ordered
