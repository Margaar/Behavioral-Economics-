    IsCodeTxt = (t as text) as logical =>
                  Text.Length(t) = 5
                  and List.AllTrue(List.Transform(Text.ToList(Text.Start(t, 2)), each _ >= "A" and _ <= "Z"))
                  and List.AllTrue(List.Transform(Text.ToList(Text.End(t, 3)),   each _ >= "0" and _ <= "9")),
    Cands     = List.Combine(List.Transform(List.FirstN(Raw, 6), (r) => List.Select(r, each _ is text))),
    Hit       = List.First(List.Select(Cands, each IsCodeTxt(Text.Trim(Text.BeforeDelimiter(_, " - ")))), null),
    IFICC     = if Hit <> null then Text.Trim(Text.BeforeDelimiter(Hit, " - ")) else fileName,







= Table.Distinct(Table.SelectColumns(CP_Data, {"IFICC", "File", "Modified"}))










let
    Sites    = { "https://generali.sharepoint.com/sites/GCSC" },
    AllFiles = Table.Combine(List.Transform(Sites, each SharePoint.Files(_, [ApiVersion = 15]))),
    Files    = Table.SelectRows(AllFiles, each
                 Text.Contains([Name], "CP-Tool 27-29")
                 and List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
                 and not Text.StartsWith([Name], "~$")),
    Res      = Table.AddColumn(Files, "Res", each try fnCPTool([Content], [Name])),
    Status   = Table.AddColumn(Res, "Status", each if [Res][HasError] then [Res][Error][Message] else "OK", type text),
    OK       = Table.SelectRows(Status, each [Status] = "OK"),
    Data     = Table.AddColumn(OK, "T", (row) =>
                 Table.AddColumn(
                   Table.AddColumn(row[Res][Value], "Modified", each row[Date modified], type datetime),
                   "Source", each row[Folder Path], type text)),
    Comb     = Table.Combine(Data[T]),

    // по каждой компании — только самый свежий файл
    MaxMod   = Table.Group(Comb, {"IFICC"}, {{"MaxMod", each List.Max([Modified]), type datetime}}),
    J        = Table.ExpandTableColumn(Table.NestedJoin(Comb, {"IFICC"}, MaxMod, {"IFICC"}, "M", JoinKind.Inner), "M", {"MaxMod"}),
    Latest   = Table.SelectRows(J, each [Modified] = [MaxMod]),
    Dedup    = Table.Distinct(Latest, {"IFICC", "Item"}),
    Out      = Table.RemoveColumns(Dedup, {"MaxMod"})
in
    Out









