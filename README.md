
====================================================================
=====================================================================

HOW TO USE
- Every query is created the same way:
  Nuova origine > Query vuota > Editor avanzato > Ctrl+A > paste/type > Fine > rename it.
- Lines starting with // are comments. You can skip them when typing.
- Keep the existing query "Workbook" (the Excel.Workbook(...) one). Don't change it.
- Delete the old queries fnGetBlock and Config / BlockConfig once the new ones work.

QUERY LIST (left panel, final state)
  Workbook            (exists)            load OFF
  fnReadSheet         function            -
  fnReadList          function            -
  BlockConfig         config table        load OFF
  CheckBlocks         debug table         load OFF
  FactData            main fact table     load ON
  DimItem             item sort order     load ON
  DimPeriod           period sort order   load ON
  List_...            optional lists      load ON (only the ones you need)

ORDER TO CREATE THEM
  1 fnReadSheet  2 BlockConfig  3 CheckBlocks  (check it!)  4 FactData
  5 DimItem  6 DimPeriod  7 fnReadList + List_ queries (optional)


=====================================================================
 1) fnReadSheet
    Reads every "label column + period columns" table on a sheet.
    It finds the tables by itself. No header texts are needed.
=====================================================================

(Sheet as table, optional FromText as nullable text, optional ToText as nullable text,
 optional SideGroup as nullable logical, optional SectionPrefix as nullable text) as table =>
let
    Clean = Table.ReplaceErrorValues(Sheet, List.Transform(Table.ColumnNames(Sheet), each {_, null})),
    Rows  = List.Buffer(Table.ToRows(Clean)),
    NR    = List.Count(Rows),
    NC    = Table.ColumnCount(Sheet),
    Txt   = (v) => if v = null then null
                   else Text.Trim(Text.Replace(Text.Replace(Text.From(v), "#(cr)", ""), "#(lf)", " ")),

    // column window: from the rightmost cell starting with FromText, up to the cell starting with ToText
    FindCols = (t as text) as list => List.Combine(List.Transform(Rows, (row) =>
                   List.Select(List.Positions(row), (c) => row{c} is text and Text.StartsWith(Txt(row{c}), t)))),
    C0 = if FromText = null or FromText = "" then 0
         else let f = FindCols(FromText) in
              if List.IsEmpty(f) then error Error.Record("FromText not found", FromText) else List.Max(f),
    C1 = if ToText = null or ToText = "" then NC - 1
         else let f = List.Select(FindCols(ToText), (c) => c > C0) in
              if List.IsEmpty(f) then error Error.Record("ToText not found", ToText) else List.Min(f) - 1,
    WinCols = List.Buffer(List.Numbers(C0, C1 - C0 + 1)),

    // period header cells: 2026, 2030*, Act 2025, For 2026, Plan 2027, HY26, Pro Forma ... 2026, Benchmark
    IsPeriod = (v) =>
        if v is number then v >= 2000 and v <= 2100 and Number.Round(v) = v
        else if v is text then
            (let s = Text.Upper(Txt(v)), d = Text.Select(s, {"0".."9"}) in
                (Text.Length(d) = 4 and Text.StartsWith(d, "20"))
                or (Text.StartsWith(s, "HY") and Text.Length(d) = 2)
                or s = "BENCHMARK")
        else false,

    IsHeader = (r as number) as logical =>
        let cells = List.Transform(WinCols, (c) => Rows{r}{c}),
            nPer  = List.Count(List.Select(cells, IsPeriod)),
            nNum  = List.Count(List.Select(cells, (v) => v is number and not IsPeriod(v)))
        in nPer >= 2 and nNum = 0,
    HdrRows = List.Buffer(List.Select(List.Numbers(0, NR), IsHeader)),

    Junk = (t) => t = null or t = "" or t = "0" or Text.StartsWith(t, "*")
                  or Text.Upper(t) = "TO ALLOCATE" or Text.StartsWith(Text.Upper(t), "WHERE NEEDED"),
    FillFwd = (lst as list) as list => List.Accumulate(lst, {}, (s, v) =>
                  s & {if v <> null and v <> "" then v else (if List.IsEmpty(s) then null else List.Last(s))}),
    Prefixes = if SectionPrefix = null or SectionPrefix = "" then {} else Text.Split(SectionPrefix, "|"),
    IsSec = (t) => t <> null and List.AnyTrue(List.Transform(Prefixes, (p) => Text.StartsWith(t, p))),

    // entity (Company #...) near each header, carried forward to the next tables
    P0Of  = (h) => List.Min(List.Select(WinCols, (c) => IsPeriod(Rows{h}{c}))),
    EntOf = (h) =>
        let p0 = P0Of(h),
            cells = List.Combine(List.Transform(List.Select(List.Numbers(h, 7, -1), (r) => r >= 0),
                        (r) => List.Transform(List.Numbers(p0 - 1, p0, -1), (c) => Txt(Rows{r}{c}))))
        in List.First(List.Select(cells, (t) => t <> null
               and Text.StartsWith(Text.Upper(t), "COMPANY")
               and not Text.StartsWith(Text.Upper(t), "COMPANIES")
               and not Text.StartsWith(Text.Upper(t), "COMPANY NAME")), null),
    EntFill = List.Buffer(FillFwd(List.Transform(HdrRows, EntOf))),

    ReadBlock = (k as number) as list =>
        let
            h       = HdrRows{k},
            hEnd    = if k + 1 < List.Count(HdrRows) then HdrRows{k + 1} else NR,
            HRow    = Rows{h},
            p0      = P0Of(h),
            BodyAll = List.Numbers(h + 1, hEnd - h - 1),
            HasText = (c) => List.AnyTrue(List.Transform(BodyAll, (r) =>
                          Rows{r}{c} is text and not Junk(Txt(Rows{r}{c})))),
            LabCand = List.Select(List.Reverse(List.Select(WinCols, (c) => c < p0)), HasText),
            LabelCol = if List.IsEmpty(LabCand) then null else LabCand{0},

            ValArea = List.Select(WinCols, (c) => c >= p0),
            HFill   = FillFwd(List.Transform(ValArea, (c) => if IsPeriod(HRow{c}) then Txt(HRow{c}) else null)),

            // optional second header row (Gross/Net, MV/Purchases/Sales...)
            SubR     = h + 1,
            SubOK    = SubR < hEnd,
            SubCells = if SubOK then List.Transform(ValArea, (c) => Rows{SubR}{c})
                       else List.Repeat({null}, List.Count(ValArea)),
            IsSubTxt = (v) => v is text and not IsPeriod(v)
                       and not List.Contains({"", "N.A.", "N.A", "-", "N/A"}, Text.Upper(Txt(v))),
            SubLab   = if SubOK and LabelCol <> null then Txt(Rows{SubR}{LabelCol}) else null,
            HasSub   = SubOK and List.Count(List.Select(SubCells, IsSubTxt)) >= 2
                       and (SubLab = null or SubLab = "" or Text.StartsWith(Text.Lower(SubLab), "amounts")),
            ValPos   = List.Select(List.Positions(ValArea), (i) =>
                           if HasSub then IsSubTxt(SubCells{i}) and HFill{i} <> null
                           else IsPeriod(HRow{ValArea{i}})),

            // optional group row above the periods (Current ... / Previous ...)
            GCells  = if h > 0 then List.Transform(ValArea, (c) =>
                          let t = Txt(Rows{h - 1}{c}) in if t = null or t = "" or IsPeriod(t) then null else t)
                      else List.Repeat({null}, List.Count(ValArea)),
            UseGrp  = List.Count(List.RemoveNulls(GCells)) >= 2,
            GFill   = if UseGrp then FillFwd(GCells) else List.Repeat({null}, List.Count(ValArea)),

            BodyStart = if HasSub then h + 2 else h + 1,
            Body    = List.Numbers(BodyStart, List.Max({0, hEnd - BodyStart})),
            UseSide = SideGroup = true and LabelCol <> null and LabelCol > 0,
            SecVals = List.Transform(Body, (r) =>
                          if UseSide then Txt(Rows{r}{LabelCol - 1})
                          else let l = Txt(Rows{r}{LabelCol}) in if IsSec(l) then l else null),
            SecFill = FillFwd(SecVals),

            HLabel  = Txt(HRow{LabelCol}),
            HLeft   = List.Select(List.Transform(List.Reverse(List.Select(WinCols, (c) => c < LabelCol)),
                          (c) => Txt(HRow{c})), (t) => t <> null and t <> "" and not IsPeriod(t)),
            FirstIt = List.First(List.Select(List.Transform(Body, (r) => Txt(Rows{r}{LabelCol})),
                          (t) => not Junk(t)), null),
            BlockNm = if HLabel <> null and HLabel <> "" and HLabel <> "0" and not IsPeriod(HLabel) then HLabel
                      else if not List.IsEmpty(HLeft) then HLeft{0}
                      else if FirstIt <> null then FirstIt
                      else "Block " & Text.From(k + 1),
            EntV    = EntFill{k},

            Recs = List.Combine(List.Transform(List.Positions(Body), (j) =>
                       let r = Body{j}, lab = Txt(Rows{r}{LabelCol}) in
                       if Junk(lab) then {}
                       else List.Transform(ValPos, (i) => [
                           BlockNo  = k + 1,
                           Block    = BlockNm,
                           Entity   = EntV,
                           Grp      = GFill{i},
                           Section  = SecFill{j},
                           RowNo    = r + 1,
                           RawItem  = lab,
                           Period   = HFill{i},
                           Measure  = if HasSub then Txt(SubCells{i}) else null,
                           RawValue = Rows{r}{ValArea{i}}
                       ])))
        in
            if LabelCol = null then {} else Recs,

    Long  = List.Combine(List.Transform(List.Positions(HdrRows), ReadBlock)),
    Cols0 = {"BlockNo", "Block", "Entity", "Grp", "Section", "RowNo", "RawItem", "Period", "Measure", "RawValue"},
    T0 = if List.IsEmpty(Long) then #table(Cols0, {}) else Table.FromRecords(Long),

    T1 = Table.AddColumn(T0, "Scenario", each
            if Text.Contains(Text.Upper((if [Grp] = null then "" else [Grp]) & " " & [Block]), "PREVIOUS")
            then "Previous" else "Current", type text),
    T2 = Table.AddColumn(T1, "Year", each
            let d = Text.Select([Period], {"0".."9"}) in
            if Text.Length(d) = 4 then Number.From(d)
            else if Text.Length(d) = 2 then 2000 + Number.From(d) else null, Int64.Type),
    T3 = Table.AddColumn(T2, "PeriodType", each
            let s = Text.Upper([Period]) in
            if Text.Contains(s, "PRO FORMA") then "Pro Forma SII"
            else if Text.StartsWith(s, "HY") then "Half-Year"
            else if Text.StartsWith(s, "ACT") then "Actual"
            else if Text.StartsWith(s, "FOR") then "Forecast"
            else if Text.StartsWith(s, "PLAN") then "Plan"
            else if s = "BENCHMARK" then "Benchmark"
            else let rest = Text.Trim(Text.Remove([Period], {"0".."9", "*"})) in
                 if rest = "" then "Year" else rest, type text),
    T4 = Table.AddColumn(T3, "SignHint", each
            let t = [RawItem] in
            if Text.Contains(t, "(+/-)") then "+/-"
            else if Text.Contains(t, "(-/+)") then "-/+"
            else if Text.Contains(t, "(-)") then "-"
            else if Text.Contains(t, "(+)") then "+" else null, type text),
    T5 = Table.AddColumn(T4, "Item", each
            Text.Trim(Text.Remove(List.Accumulate({"(+/-)", "(-/+)", "(-)", "(+)"}, [RawItem],
                (s, p) => Text.Replace(s, p, "")), {"*"})), type text),
    T6 = Table.AddColumn(T5, "IsTotal", each Text.StartsWith(Text.Upper([Item]), "TOTAL"), type logical),
    T7 = Table.AddColumn(T6, "Value", each
            if [RawValue] is number then [RawValue] else try Number.From([RawValue]) otherwise null, type number),
    T8 = Table.RenameColumns(T7, {{"Grp", "Group"}}),
    Result = Table.SelectColumns(T8, {"BlockNo", "Block", "Entity", "Group", "Scenario", "Section", "RowNo",
                 "Item", "SignHint", "IsTotal", "Period", "PeriodType", "Year", "Measure", "Value"})
in
    Result


=====================================================================
 2) BlockConfig          (load OFF)
    One line per sheet. FromText = a text that marks where the INPUT
    table starts (rightmost match is used). ToText = where to stop.
    SideGroup = true only for ORMT Details. SectionPrefix = row labels
    that open a sub-section (separated by |).
    Check that the Sheet names match your tab names exactly.
=====================================================================

let
    Config = #table(
        type table [Topic = text, Sheet = text, FromText = nullable text, ToText = nullable text,
                    SideGroup = logical, SectionPrefix = nullable text],
        {
            {"Remittance",            "Total Remittance",              "Instructions",                      null, false, null},
            {"Remittance by Company", "Remittance by Company",         "Companies in Full and Light Scope", null, false, null},
            {"Remittance Holding",    "Remittance Holding",            "Cash Position",                     null, false, null},
            {"Capital Needs",         "Capital Needs",                 null, "Capital Needs (Inflows)",           false, null},
            {"Solvency",              "Solvency, Distributions & FTC", "Instructions",                      null, false, null},
            {"Solvency AoM",          "Solvency AoM",                  "EOF (BoP)",                         null, false, "EOF (BoP)|SCR (BoP)"},
            {"Free Tangible Capital", "Free Tangible Capital",         "Companies in Full Scope",           null, false, null},
            {"Sensitivities",         "Sens on Net Result & Solvency", "Companies in Full Scope",           null, false, "Impact of"},
            {"SAA",                   "SAA",                           "Cash & Cash Equivalent", "Sales and Redemptions",  false, null},
            {"Life KPIs",             "Life KPIs",                     "NBV",                               null, false, null},
            {"Non-Life KPIs",         "Non-Life KPIs",                 "Earned Premiums",                   null, false, "Earned Premiums|Non-life SII|CoR"},
            {"HMS Impacts",           "HMS impacts",                   "Instructions",                      null, false, null},
            {"Model Changes",         "Model Changes Details",         "Instructions",                      null, false, null},
            {"ORMT Details",          "ORMT Details",                  "ORMT Impacts",                      null, true,  null}
        })
in
    Config


=====================================================================
 3) CheckBlocks          (load OFF)  - open this FIRST after changes
=====================================================================

let
    Checked = Table.AddColumn(BlockConfig, "Status", each
        let r = try fnReadSheet(Workbook{[Item = [Sheet], Kind = "Sheet"]}[Data],
                                [FromText], [ToText], [SideGroup], [SectionPrefix])
        in if r[HasError] then "ERROR: " & r[Error][Reason] & " - " & Text.From(r[Error][Message])
           else Text.From(Table.RowCount(r[Value])) & " rows | "
                & Text.From(List.Count(List.Distinct(r[Value][BlockNo]))) & " blocks: "
                & Text.Combine(List.Distinct(List.Transform(r[Value][Block], Text.From)), " / ")),
    Result = Table.SelectColumns(Checked, {"Topic", "Sheet", "Status"})
in
    Result


=====================================================================
 4) FactData             (load ON)  - the one table for all topics
=====================================================================

let
    WithData = Table.AddColumn(BlockConfig, "Data", each
        try fnReadSheet(Workbook{[Item = [Sheet], Kind = "Sheet"]}[Data],
                        [FromText], [ToText], [SideGroup], [SectionPrefix])
        otherwise null),
    Found    = Table.SelectRows(WithData, each [Data] <> null),
    Expanded = Table.ExpandTableColumn(Found, "Data",
                   {"BlockNo", "Block", "Entity", "Group", "Scenario", "Section", "RowNo",
                    "Item", "SignHint", "IsTotal", "Period", "PeriodType", "Year", "Measure", "Value"}),
    Removed  = Table.RemoveColumns(Expanded, {"FromText", "ToText", "SideGroup", "SectionPrefix"}),
    Typed    = Table.TransformColumnTypes(Removed, {
                   {"BlockNo", Int64.Type}, {"RowNo", Int64.Type}, {"Year", Int64.Type},
                   {"Value", type number}, {"IsTotal", type logical},
                   {"Block", type text}, {"Entity", type text}, {"Group", type text},
                   {"Section", type text}, {"Item", type text}, {"Period", type text},
                   {"Measure", type text}})
in
    Typed


=====================================================================
 5) DimItem              (load ON)  - keeps the Excel row order
=====================================================================

let
    Src     = Table.SelectColumns(FactData, {"Item", "RowNo"}),
    Grouped = Table.Group(Src, {"Item"}, {{"ItemOrder", each List.Min([RowNo]), Int64.Type}},
                          GroupKind.Global, Comparer.OrdinalIgnoreCase)
in
    Grouped


=====================================================================
 6) DimPeriod            (load ON)  - Act 2025 < HY26 < For 2026 < Plan 2027 ...
=====================================================================

let
    Src     = Table.SelectColumns(FactData, {"Period", "Year", "PeriodType"}),
    Grouped = Table.Group(Src, {"Period"},
                  {{"Year", each List.Max([Year]), Int64.Type},
                   {"PeriodType", each List.First([PeriodType]), type text}},
                  GroupKind.Global, Comparer.OrdinalIgnoreCase),
    Ordered = Table.AddColumn(Grouped, "PeriodOrder", each
                  (if [Year] = null then 9999 else [Year]) * 10
                  + (if [PeriodType] = "Actual" then 1
                     else if [PeriodType] = "Half-Year" then 2
                     else if [PeriodType] = "Pro Forma SII" then 3
                     else if [PeriodType] = "Forecast" then 4
                     else if [PeriodType] = "Year" then 5
                     else if [PeriodType] = "Plan" then 6 else 7), Int64.Type)
in
    Ordered


=====================================================================
 7) fnReadList  (OPTIONAL - only for the record lists:
    Debt Management, Options Monitoring, ORMT, Local RAF, Initiatives)
    KeyText    = a text in the TOP header row of the table
    HeaderRows = how many rows the header has
    Pick       = 1 first table found, 2 second, 0 last
    MaxRows    = how many template rows to read below the header
    FillDownFirst = copy the first N columns down (merged cells)
=====================================================================

(Sheet as table, FromText as nullable text, ToText as nullable text, KeyText as text,
 HeaderRows as number, Pick as number, MaxRows as number, optional FillDownFirst as nullable number) as table =>
let
    Clean = Table.ReplaceErrorValues(Sheet, List.Transform(Table.ColumnNames(Sheet), each {_, null})),
    Rows  = List.Buffer(Table.ToRows(Clean)),
    NR    = List.Count(Rows),
    NC    = Table.ColumnCount(Sheet),
    Txt   = (v) => if v = null then null
                   else Text.Trim(Text.Replace(Text.Replace(Text.From(v), "#(cr)", ""), "#(lf)", " ")),
    FindCols = (t as text) as list => List.Combine(List.Transform(Rows, (row) =>
                   List.Select(List.Positions(row), (c) => row{c} is text and Text.StartsWith(Txt(row{c}), t)))),
    C0 = if FromText = null or FromText = "" then 0
         else let f = FindCols(FromText) in
              if List.IsEmpty(f) then error Error.Record("FromText not found", FromText) else List.Max(f),
    C1 = if ToText = null or ToText = "" then NC - 1
         else let f = List.Select(FindCols(ToText), (c) => c > C0) in
              if List.IsEmpty(f) then error Error.Record("ToText not found", ToText) else List.Min(f) - 1,
    WinCols = List.Numbers(C0, C1 - C0 + 1),

    Hits = List.Select(List.Numbers(0, NR), (r) => List.AnyTrue(List.Transform(WinCols, (c) =>
               let t = Txt(Rows{r}{c}) in t <> null and Text.StartsWith(t, KeyText)))),
    R  = if List.IsEmpty(Hits) or List.Count(Hits) < Pick then error Error.Record("KeyText not found", KeyText)
         else if Pick = 0 then List.Last(Hits) else Hits{Pick - 1},
    HR = List.Numbers(R, HeaderRows),

    FFRow = (r as number, ff as logical) as list => List.Accumulate(WinCols, {}, (s, c) =>
                let t = Txt(Rows{r}{c}),
                    v = if t <> null and t <> "" then t
                        else if ff and not List.IsEmpty(s) then List.Last(s) else null
                in s & {v}),
    HdrGrid  = List.Transform(List.Positions(HR), (i) => FFRow(HR{i}, i < HeaderRows - 1)),
    Orig     = List.Transform(HR, (r) => List.Transform(WinCols, (c) => let t = Txt(Rows{r}{c}) in t <> null and t <> "")),
    KeepIdx  = List.Select(List.Positions(WinCols), (j) => List.AnyTrue(List.Transform(Orig, (o) => o{j}))),
    RawNames = List.Transform(KeepIdx, (j) =>
                   Text.Combine(List.Distinct(List.RemoveNulls(List.Transform(HdrGrid, (g) => g{j}))), " - ")),
    Names    = List.Accumulate(RawNames, {}, (s, n) =>
                   s & {if List.Contains(s, n) then n & " (" & Text.From(List.Count(s) + 1) & ")" else n}),
    KeepCols = List.Transform(KeepIdx, (j) => WinCols{j}),
    LabelC   = List.Min(KeepCols) - 1,
    AllCols  = if LabelC >= 0 then {LabelC} & KeepCols else KeepCols,
    AllNames = if LabelC >= 0 then {"Label"} & Names else Names,

    First = List.Last(HR) + 1,
    BodyR = List.Numbers(First, List.Max({0, List.Min({MaxRows, NR - First})})),
    T0 = Table.FromRows(List.Transform(BodyR, (r) => List.Transform(AllCols, (c) => Rows{r}{c})), AllNames),
    NFill = if FillDownFirst = null then 0 else FillDownFirst,
    T1 = if NFill = 0 then T0 else Table.FillDown(T0, List.FirstN(Names, NFill)),
    DataCols = List.Skip(Names, NFill),
    Meaningful = (v) => v <> null and not (v is text and Text.Trim(v) = "") and not (v is number and v = 0),
    T2 = Table.AddIndexColumn(T1, "RowNo", 1, 1, Int64.Type),
    T3 = Table.SelectRows(T2, (row) =>
             List.Count(List.Select(List.Transform(DataCols, (n) => Record.Field(row, n)), Meaningful)) > 0),
    T4 = Table.SelectRows(T3, (row) => not List.AnyTrue(List.Transform(AllNames, (n) =>
             let v = Record.Field(row, n) in v is text and Text.StartsWith(Text.Trim(v), KeyText)))),
    T5 = if List.Contains(AllNames, "Label") and List.NonNullCount(Table.Column(T4, "Label")) = 0
         then Table.RemoveColumns(T4, {"Label"}) else T4
in
    T5


=====================================================================
 8) List queries (OPTIONAL, one query each, load ON). Create only
    the ones you need. Each is 3 lines.
=====================================================================

--- List_DebtExisting
let
    Sheet  = Workbook{[Item = "Debt Management", Kind = "Sheet"]}[Data],
    Result = fnReadList(Sheet, "Proposals on Existing Debt", null, "Borrower", 1, 1, 37, 0)
in
    Result

--- List_DebtNew
let
    Sheet  = Workbook{[Item = "Debt Management", Kind = "Sheet"]}[Data],
    Result = fnReadList(Sheet, "Proposals on Existing Debt", null, "Borrower", 1, 2, 35, 0)
in
    Result

--- List_Options
let
    Sheet  = Workbook{[Item = "Options Monitoring", Kind = "Sheet"]}[Data],
    Result = fnReadList(Sheet, "ID", null, "Option Buyer", 2, 1, 15, 0)
in
    Result

--- List_ORMT
let
    Sheet  = Workbook{[Item = "ORMT", Kind = "Sheet"]}[Data],
    Result = fnReadList(Sheet, "ID", null, "Cedant", 1, 1, 50, 0)
in
    Result

--- List_LocalRAF
let
    Sheet  = Workbook{[Item = "Local RAF", Kind = "Sheet"]}[Data],
    Result = fnReadList(Sheet, "Currently in place", null, "Currently in place", 3, 1, 30, 0)
in
    Result

--- List_Initiatives
let
    Sheet  = Workbook{[Item = "Initiatives", Kind = "Sheet"]}[Data],
    Result = fnReadList(Sheet, "ID", null, "Expected Timeline", 2, 1, 100, 8)
in
    Result


=====================================================================
 9) MODEL + DAX (after Chiudi e applica)
=====================================================================

Model view:
  DimItem[Item]     1 -> *  FactData[Item]
  DimPeriod[Period] 1 -> *  FactData[Period]
Data view:
  DimItem   > Item   > Ordina per colonna > ItemOrder
  DimPeriod > Period > Ordina per colonna > PeriodOrder
  FactData  > Year, RowNo, BlockNo > Riepilogo: Non riepilogare

Measures (Home > Nuova misura, one at a time):

Value = SUM ( FactData[Value] )

Value Current = CALCULATE ( [Value], FactData[Scenario] = "Current" )

Value Previous = CALCULATE ( [Value], FactData[Scenario] = "Previous" )

Delta = [Value Current] - [Value Previous]

Delta % = DIVIDE ( [Delta], ABS ( [Value Previous] ) )

Value excl Totals = CALCULATE ( [Value], FactData[IsTotal] = FALSE () )
