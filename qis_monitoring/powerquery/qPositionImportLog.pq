// qPositionImportLog: フォルダ内の全ポジションファイルの取込診断（エラー・重複・基準日検証）
// 読込先: 新規シート（例: 05_Position_Raw の右側や別シート）。毎週の取込前に確認する。
let
    Params     = Excel.CurrentWorkbook(){[Name = "tblPQParams"]}[Content],
    P          = (name as text) => Params{[Param = name]}[Value],
    Files      = Folder.Files(P("PositionFolder")),
    Candidates = Table.SelectRows(Files, each Text.StartsWith(Text.Lower([Name]), "positions_")
                                         and List.Contains({".csv", ".xlsx"}, Text.Lower([Extension]))),
    Diag       = Table.AddColumn(Candidates, "Diag", each
                    let
                        r = try fnReadPositionFile([Content], [Extension], P("PositionSheetName"), [Name])
                    in
                        if r[HasError] then
                            [Status = "ERROR", Message = r[Error][Message], Rows = null, PosDates = null, Indices = null, DupRows = null, NullTicker = null, BadDate = null]
                        else
                            let
                                t        = r[Value],
                                dates    = List.Distinct(List.RemoveNulls(t[Position Date])),
                                keyCols  = {"Position Date", "Index Name", "BloombergTicker", "open_date", "Notional(USD)"},
                                dup      = Table.RowCount(t) - Table.RowCount(Table.Distinct(t, keyCols)),
                                nullTk   = Table.RowCount(Table.SelectRows(t, each [BloombergTicker] = null or [BloombergTicker] = "")),
                                badDate  = Table.RowCount(Table.SelectRows(t, each [Position Date] = null or [Maturity Date] = null))
                            in
                                [Status  = if List.Count(dates) <> 1 or dup > 0 or nullTk > 0 or badDate > 0 then "WARNING" else "OK",
                                 Message = Text.Combine(List.RemoveNulls({
                                              if List.Count(dates) <> 1 then "Position Date が複数/欠損" else null,
                                              if dup > 0 then "重複行 " & Text.From(dup) else null,
                                              if nullTk > 0 then "Ticker欠損 " & Text.From(nullTk) else null,
                                              if badDate > 0 then "日付変換不能 " & Text.From(badDate) else null}), " / "),
                                 Rows = Table.RowCount(t), PosDates = Text.Combine(List.Transform(dates, Date.ToText), ","),
                                 Indices = Text.Combine(List.Distinct(t[Index Name]), " | "), DupRows = dup, NullTicker = nullTk, BadDate = badDate]),
    Expanded   = Table.ExpandRecordColumn(Table.SelectColumns(Diag, {"Name", "Date modified", "Diag"}), "Diag",
                    {"Status", "Message", "Rows", "PosDates", "Indices", "DupRows", "NullTicker", "BadDate"})
in
    Expanded
