# Behavioral-Economics-
Beh economics assignment 


(Sheet as table, HeaderText as text, Scenario as text, optional UseRightmost as logical) as table =>
let
    Clean = Table.ReplaceErrorValues(Sheet, List.Transform(Table.ColumnNames(Sheet), each {_, null})),
    Rows  = Table.ToRows(Clean),

    IsPrefix = Text.EndsWith(HeaderText, "*"),
    Key      = if IsPrefix then Text.Start(HeaderText, Text.Length(HeaderText) - 1) else HeaderText,
    Match    = (v) => v <> null and
                 (let t = Text.Trim(Text.From(v)) in
                  if IsPrefix then Text.StartsWith(t, Key) else t = Key),

    Hits = List.Combine(
        List.Transform(List.Positions(Rows), (r) =>
            let row = Rows{r} in
            List.Transform(List.Select(List.Positions(row), (c) => Match(row{c})),
                           (c) => [R = r, C = c]))),

    Chosen =
        if List.IsEmpty(Hits) then error Error.Record("Block not found", HeaderText)
        else if IsPrefix then Hits
        else if UseRightmost = true then {List.Last(List.Sort(Hits, (a, b) => Value.Compare(a[C], b[C])))}
        else {List.First(Hits)},

    IsPeriod = (v) => v <> null and
        (let s = Text.Upper(Text.Trim(Text.From(v))),
             t = Text.Select(s, {"0".."9"})
         in (Text.Length(t) = 4 and Text.StartsWith(t, "20"))
            or (Text.Length(t) = 2 and Text.StartsWith(s, "HY"))),

    ReadBlock = (hit as record) as list =>
        let
            HeaderRow   = Rows{hit[R]},
            After       = List.Skip(List.Positions(HeaderRow), hit[C] + 1),
            AfterNN     = List.Skip(After, (c) => HeaderRow{c} = null),
            PerPos      = List.FirstN(AfterNN, (c) => IsPeriod(HeaderRow{c})),
            IsEmptyRow  = (row) => List.IsEmpty(List.RemoveNulls(
                              List.Transform({hit[C]} & PerPos, (c) => row{c}))),
            IsHeaderRow = (row) => List.AnyTrue(
                              List.Transform(PerPos, (c) => row{c} is text and IsPeriod(row{c}))),
            Body        = List.FirstN(List.Skip(Rows, hit[R] + 1),
                              (row) => not IsEmptyRow(row) and not IsHeaderRow(row)),
            BlockName   = Text.Trim(Text.From(HeaderRow{hit[C]}))
        in
            List.Combine(List.Transform(Body, (row) =>
                List.Transform(PerPos, (c) => [
                    Block    = BlockName,
                    RawItem  = if row{hit[C]} = null then "" else Text.Trim(Text.From(row{hit[C]})),
                    Period   = Text.Trim(Text.From(HeaderRow{c})),
                    RawValue = row{c}
                ]))),

    Long = List.Combine(List.Transform(Chosen, ReadBlock)),
    T0 = if List.IsEmpty(Long)
         then #table({"Block", "RawItem", "Period", "RawValue"}, {})
         else Table.FromRecords(Long),
    T1 = Table.SelectRows(T0, each [RawItem] <> "" and [RawItem] <> "0"),
    T2 = Table.AddColumn(T1, "Scenario", each Scenario, type text),
    T3 = Table.AddColumn(T2, "Year", each
            let t = Text.Select([Period], {"0".."9"}) in
            if Text.Length(t) = 2 then 2000 + Number.From(t) else Number.From(t), Int64.Type),
    T4 = Table.AddColumn(T3, "PeriodType", each
            let s = Text.Upper([Period]) in
            if Text.Contains(s, "PRO FORMA") then "Pro Forma SII"
            else if Text.StartsWith(s, "HY") then "Half-Year"
            else if Text.StartsWith(s, "ACT") then "Actual"
            else if Text.StartsWith(s, "FOR") then "Forecast"
            else if Text.StartsWith(s, "PLAN") then "Plan"
            else "Year", type text),
    T5 = Table.AddColumn(T4, "SignHint", each
            if Text.Contains([RawItem], "(-/+)") then "-/+"
            else if Text.Contains([RawItem], "(-)") then "-"
            else if Text.Contains([RawItem], "(+)") then "+"
            else null, type text),
    T6 = Table.AddColumn(T5, "Item", each
            Text.Trim(Text.Replace(Text.Replace(Text.Replace([RawItem], "(-/+)", ""), "(-)", ""), "(+)", "")),
            type text),
    T7 = Table.AddColumn(T6, "IsTotal", each Text.StartsWith([Item], "Total"), type logical),
    T8 = Table.AddColumn(T7, "Value", each
            if [RawValue] = null then 0
            else try Number.From([RawValue]) otherwise 0, type number),
    Result = Table.SelectColumns(T8,
            {"Scenario", "Block", "Item", "Period", "PeriodType", "Year", "SignHint", "IsTotal", "Value"})
in
    Result






    let
    WithData = Table.AddColumn(BlockConfig, "Data", each
        try fnGetBlock(
                Workbook{[Item = [Sheet], Kind = "Sheet"]}[Data],
                [HeaderText], [Scenario], [UseRightmost])
        otherwise null),
    Found    = Table.SelectRows(WithData, each [Data] <> null),
    Expanded = Table.ExpandTableColumn(Found, "Data",
                   {"Block", "Item", "Period", "PeriodType", "Year", "SignHint", "IsTotal", "Value"}),
    NoTotals = Table.SelectRows(Expanded, each [IsTotal] = false),
    Result   = Table.SelectColumns(NoTotals,
                   {"Topic", "Sheet", "Block", "Scenario", "Item", "Period", "PeriodType", "Year", "SignHint", "Value"}),
    Typed    = Table.TransformColumnTypes(Result, {{"Year", Int64.Type}, {"Value", type number}})
in
    Typed





            {
            {"Remittance",    "Total Remittance",              "Current Finalized/Planned Remittance",    "Current",  true},
            {"Remittance",    "Total Remittance",              "Previous Planned Remittance",             "Previous", true},
            {"Capital Needs", "Capital Needs",                 "Current Finalized/Planned Capital Needs", "Current",  false},
            {"Capital Needs", "Capital Needs",                 "Previous Planned Capital Needs",          "Previous", false},
            {"Solvency",      "Solvency, Distributions & FTC", "Company #*",                              "Current",  false}
        })





        
