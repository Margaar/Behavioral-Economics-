let
    // ← сюда все сайты SharePoint (только корень: до /sites/NOME)
    Sites    = {
        "https://generali.sharepoint.com/sites/GCSC"
        // , "https://generali.sharepoint.com/sites/ALTRO_SITO"
    },
    AllFiles = Table.Combine(List.Transform(Sites, each SharePoint.Files(_, [ApiVersion = 15]))),
    Files    = Table.SelectRows(AllFiles, each
                 Text.Contains([Name], "CP-Tool 27-29")
                 and List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
                 and not Text.StartsWith([Name], "~$")),

    // код BU из имени файла: "DAS - CP-Tool..." / "GIT_CP-Tool..." -> DAS / GIT
    WithBU   = Table.AddColumn(Files, "BU", each
                 Text.Trim(Text.BeforeDelimiter(Text.Replace([Name], "_CP-Tool", " - CP-Tool"), " - CP-Tool")),
                 type text),

    // если у BU несколько версий файла — берём самую свежую
    Sorted   = Table.Buffer(Table.Sort(WithBU, {{"Date modified", Order.Descending}})),
    Latest   = Table.Distinct(Sorted, {"BU"}),

    Res      = Table.AddColumn(Latest, "Res", each try fnCPTool([Content], [Name])),
    Status   = Table.AddColumn(Res, "Status", each if [Res][HasError] then [Res][Error][Message] else "OK", type text),
    OK       = Table.SelectRows(Status, each [Status] = "OK"),
    Data     = Table.AddColumn(OK, "T", (row) =>
                 let t = row[Res][Value]
                 in  Table.AddColumn(
                       Table.AddColumn(t, "BU", each row[BU], type text),
                       "Source", each row[Folder Path], type text)),
    Comb     = Table.Combine(Data[T])
in
    Comb
