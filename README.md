let
    // --- источник: выбери ОДИН вариант ---
    // A) файл на компьютере / сетевом диске:
    Src     = Excel.Workbook(File.Contents("C:\...\ИМЯ_ФАЙЛА.xlsx"), null, true),

    // B) файл на SharePoint (закомментируй A, раскомментируй B):
    // Src  = Excel.Workbook(
    //          Table.SelectRows(
    //              SharePoint.Files("https://generali.sharepoint.com/sites/GCSC", [ApiVersion = 15]),
    //              each [Name] = "ИМЯ_ФАЙЛА.xlsx"){0}[Content], null, true),

    Sh      = Src{[Item = "7.4 Data Export", Kind = "Sheet"]}[Data],

    // находим строку заголовков, где бы она ни была
    Rows    = Table.ToRows(Sh),
    HdrI    = List.PositionOf(
                List.Transform(Rows, (r) => List.Contains(
                    List.Transform(r, each try Text.Trim(Text.From(_)) otherwise null),
                    "Legal Entity Code")),
                true),
    Tbl     = Table.PromoteHeaders(Table.Skip(Sh, HdrI), [PromoteAllScalars = true]),

    Keep    = Table.SelectColumns(Tbl, {"Legal Entity Code", "Legal Entity Name"}),
    Clean   = Table.TransformColumns(Keep, {
                {"Legal Entity Code", each Text.Upper(Text.Trim(Text.From(_))), type text},
                {"Legal Entity Name", each Text.Trim(Text.From(_)), type text}}),
    NoBlank = Table.SelectRows(Clean, each [Legal Entity Code] <> null and [Legal Entity Code] <> ""),
    Dist    = Table.Distinct(NoBlank, {"Legal Entity Code"})
in
    Dist
