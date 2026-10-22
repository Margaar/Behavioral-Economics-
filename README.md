let
    Src = Folder.Files(FolderPath),
    Xl  = Table.SelectRows(Src, each
              List.Contains({".xlsx", ".xlsm", ".csv"}, Text.Lower([Extension]))
              and not Text.StartsWith([Name], "~$"))
in
    Xl






    let
    F     = Table.SelectRows(Files, each Text.Contains([Name], "CMP")),
    Bin   = F{0}[Content],
    C1    = Csv.Document(Bin, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Semi  = Table.ColumnCount(C1) < 5,
    Sh    = if Semi then Csv.Document(Bin, [Delimiter = ";", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]) else C1,
    Prom  = Table.PromoteHeaders(Sh, [PromoteAllScalars = true]),
    Std   = {"Plan", "Version", "Time", "LECode", "LEName", "Segment", "Process", "AccountCode",
             "Description", "Counterparty", "InjectionType", "CmpScenario", "LocalCcy",
             "ValueLocal", "ReportingCcy", "FXRate", "ValueEUR"},
    Old   = List.FirstN(Table.ColumnNames(Prom), List.Count(Std)),
    Ren   = Table.RenameColumns(Prom, List.Zip({Old, List.FirstN(Std, List.Count(Old))})),
    Typed = Table.TransformColumnTypes(Ren, {{"ValueEUR", type number}, {"ValueLocal", type number}},
                if Semi then "it-IT" else "en-US")
in
    Typed
