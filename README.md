let
    Sheet = Workbook{[Item = "Debt Management", Kind = "Sheet"]}[Data],
    Map = (n as text) as text =>
        let u = Text.Upper(n) in
        if Text.StartsWith(u, "ID") then "ID"
        else if Text.StartsWith(u, "BORROWER") then "Borrower"
        else if Text.StartsWith(u, "LENDER") then "Lender"
        else if Text.Contains(u, "MATURITY") and not Text.Contains(u, "EXPECTED") then "Maturity"
        else if Text.Contains(u, "CALL DATE") then "CallDate"
        else if Text.StartsWith(u, "TIERING") then "Tiering"
        else if Text.StartsWith(u, "CURRENCY") then "Currency"
        else if Text.Contains(u, "TRANCHE NUMBER") then "Tranche"
        else if Text.Contains(u, "UPDATED") then "PrincipalAfter"
        else if Text.Contains(u, "EXPECTED MONTH") then "ExpectedDate"
        else if Text.Contains(u, "PRINCIPAL AMOUNT AT") and Text.Contains(u, "EUR") then "PrincipalEUR"
        else if Text.Contains(u, "PRINCIPAL AMOUNT AT") then "PrincipalCCY"
        else if Text.Contains(u, "TRANCHE") and Text.Contains(u, "EUR") then "TrancheEUR"
        else if Text.Contains(u, "TRANCHE") then "TrancheCCY"
        else n,
    Std = (t as table, typ as text) as table =>
        let old  = Table.ColumnNames(t),
            new  = List.Accumulate(old, {}, (s, n) =>
                       let m = Map(n) in
                       s & { if not List.Contains(s, m) then m
                             else if not List.Contains(s, n) then n
                             else n & " #" & Text.From(List.Count(s)) }),
            r    = Table.RenameColumns(t, List.Zip({old, new}))
        in Table.AddColumn(r, "Type", each typ, type text),
    Ex  = Std(fnReadList(Sheet, "Proposals on Existing Debt", null, "Borrower", 1, 1, 37, 0), "Existing"),
    Nw  = Std(fnReadList(Sheet, "Proposals on Existing Debt", null, "Borrower", 1, 2, 35, 0), "New"),
    All = Table.Combine({Ex, Nw}),
    ToNum = (v) => if v is number then v else try Number.From(v) otherwise null,
    Num = Table.TransformColumns(All,
              {{"PrincipalEUR", ToNum, type number}, {"TrancheEUR", ToNum, type number},
               {"PrincipalAfter", ToNum, type number}, {"PrincipalCCY", ToNum, type number},
               {"TrancheCCY", ToNum, type number}}, null, MissingField.Ignore),
    Yr  = Table.AddColumn(Num, "Year", each
              let d = try Record.Field(_, "ExpectedDate") otherwise null in
              if d is date then Date.Year(d)
              else if d is datetime then Date.Year(DateTime.Date(d))
              else if d is number then Date.Year(Date.From(d))
              else try Date.Year(Date.FromText(Text.From(d), [Culture = "en-GB"])) otherwise null,
              Int64.Type)
in
    Yr
