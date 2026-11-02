// qPositionAll: フォルダ内の全ポジションファイルを結合（履歴の初期移行・突合用）。
// 通常運用の履歴保存は QIS_ArchivePositions（値で保存）を使う。元ファイル更新による履歴の変化を避けるため。
let
    Params     = Excel.CurrentWorkbook(){[Name = "tblPQParams"]}[Content],
    P          = (name as text) => Params{[Param = name]}[Value],
    Files      = Folder.Files(P("ArchiveFolder")),
    Candidates = Table.SelectRows(Files, each Text.StartsWith(Text.Lower([Name]), "positions_")
                                         and List.Contains({".csv", ".xlsx"}, Text.Lower([Extension]))),
    Loaded     = Table.AddColumn(Candidates, "Data", each fnReadPositionFile([Content], [Extension], P("PositionSheetName"), [Name])),
    Combined   = Table.Combine(Loaded[Data]),
    Snapshot   = Table.AddColumn(Combined, "Snapshot ID", each Date.ToText([Position Date], "yyyyMMdd") & "|" & [Index Name] & "|" & [Source File], type text),
    // 同一ファイルを複数回置いても 1 回分だけ残す
    Dedup      = Table.Distinct(Snapshot, {"Snapshot ID", "BloombergTicker", "open_date", "Notional(USD)", "Strike", "Maturity Date"})
in
    Dedup
