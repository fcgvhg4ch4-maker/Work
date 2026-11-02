// fnReadPositionFile: 1 ファイル（CSV / xlsx）を読み、tblColMap に従って列名・型を標準化する
// 引数: content = Folder.Files の [Content]、ext = 拡張子、sheetName = xlsx のシート名、fileName = ファイル名
(content as binary, ext as text, sheetName as text, fileName as text) as table =>
let
    ColMap   = Excel.CurrentWorkbook(){[Name = "tblColMap"]}[Content],
    Raw      = if Text.Lower(ext) = ".csv" then
                   Table.PromoteHeaders(Csv.Document(content, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]), [PromoteAllScalars = true])
               else
                   let
                       Book  = Excel.Workbook(content, null, true),
                       Sheet = try Book{[Item = sheetName, Kind = "Sheet"]}[Data] otherwise Book{0}[Data]
                   in
                       Table.PromoteHeaders(Sheet, [PromoteAllScalars = true]),
    Trimmed  = Table.TransformColumnNames(Raw, Text.Trim),
    // 必須列の存在チェック（欠けている場合はエラーで止める＝黙って空列にしない）
    Required = Table.SelectRows(ColMap, each [Required] = "Y")[SourceColumn],
    Missing  = List.Difference(Required, Table.ColumnNames(Trimmed)),
    Checked  = if List.Count(Missing) > 0
               then error Error.Record("MissingColumn", fileName & " に必須列がありません: " & Text.Combine(Missing, ", "))
               else Trimmed,
    // 列名の標準化（SourceColumn → StandardColumn）。マップにない列は除外
    Pairs    = List.Select(List.Zip({ColMap[SourceColumn], ColMap[StandardColumn]}), each List.Contains(Table.ColumnNames(Checked), _{0})),
    Selected = Table.SelectColumns(Checked, List.Transform(Pairs, each _{0})),
    Renamed  = Table.RenameColumns(Selected, Pairs),
    // 型変換（日付は yyyy-MM-dd / yyyy/MM/dd / Excel シリアル に対応。変換不能はエラー値→品質チェックで検出）
    ToDate   = (x) => if x = null or x = "" then null
                      else if Value.Is(x, type date) then x
                      else if Value.Is(x, type datetime) then Date.From(x)
                      else if Value.Is(x, type number) then Date.From(x)
                      else Date.FromText(Text.Replace(Text.From(x), "/", "-"), [Format = "yyyy-MM-dd"]),
    ToNum    = (x) => if x = null or x = "" then null else Number.From(x, "en-US"),
    TypeOf   = (std as text) => ColMap{[StandardColumn = std]}[DataType],
    Typed    = Table.TransformColumns(Renamed,
                   List.Transform(Table.ColumnNames(Renamed), (c) =>
                       {c, if TypeOf(c) = "date" then ToDate else if TypeOf(c) = "number" then ToNum else each if _ = null then null else Text.Trim(Text.From(_))})),
    WithSrc  = Table.AddColumn(Typed, "Source File", each fileName, type text)
in
    WithSrc
