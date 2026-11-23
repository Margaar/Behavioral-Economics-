let
    Sites    = {
        "https://generali.sharepoint.com/sites/GCSC"
        // , "https://generali.sharepoint.com/sites/ALTRO_SITO"
    },
    AllFiles = Table.Combine(List.Transform(Sites, each SharePoint.Files(_, [ApiVersion = 15]))),
    Files    = Table.SelectRows(AllFiles, each
                 Text.Contains([Name], "CP-Tool 27-29")
                 and List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
                 and not Text.StartsWith([Name], "~$")),

    // IFICC из имени файла: ищем кусок вида XX999 (FR014, AT004, HU002...)
    GetCode  = (name as text) as nullable text =>
        let
            base   = Text.BeforeDelimiter(name, ".", {0, RelativePosition.FromEnd}),
            parts  = Text.SplitAny(base, "_ -"),
            isCode = (t as text) as logical =>
                        Text.Length(t) = 5
                        and List.AllTrue(List.Transform(Text.ToList(Text.Start(t, 2)), each _ >= "A" and _ <= "Z"))
                        and List.AllTrue(List.Transform(Text.ToList(Text.End(t, 3)),   each _ >= "0" and _ <= "9"))
        in
            List.First(List.Select(parts, isCode), null),

    WithCode = Table.AddColumn(Files, "Code", each GetCode([Name]), type text),
    Known    = Table.SelectRows(WithCode, each [Code] <> null),   // сводные ALL/DAS/... без кода исключаются

    // если у компании несколько файлов — берём самый свежий
    Sorted   = Table.Buffer(Table.Sort(Known, {{"Date modified", Order.Descending}})),
    Latest   = Table.Distinct(Sorted, {"Code"}),

    Res      = Table.AddColumn(Latest, "Res", each try fnCPTool([Content], [Name])),
    Status   = Table.AddColumn(Res, "Status", each if [Res][HasError] then [Res][Error][Message] else "OK", type text),
    OK       = Table.SelectRows(Status, each [Status] = "OK"),
    Data     = Table.AddColumn(OK, "T", (row) =>
                 let t = Table.RemoveColumns(row[Res][Value], {"IFICC"})
                 in  Table.AddColumn(
                       Table.AddColumn(t, "IFICC", each row[Code], type text),
                       "Source", each row[Folder Path], type text)),
    Comb     = Table.Combine(Data[T])
in
    Comb
