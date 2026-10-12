(Sheet as table, optional FromText as nullable text, optional ToText as nullable text,
 optional SideGroup as nullable logical, optional SectionPrefix as nullable text) as table =>
let
    Clean = Table.ReplaceErrorValues(Sheet, List.Transform(Table.ColumnNames(Sheet), each {_, null})),
    Rows  = List.Buffer(Table.ToRows(Clean)),
    NR    = List.Count(Rows),
    NC    = Table.ColumnCount(Sheet),
    EntNames = List.Buffer(List.Transform(List.RemoveNulls(EntityList[Entity]), each Text.Upper(Text.Trim(_)))),
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
    WinCols = List.Buffer(List.Numbers(C0, C1 - C0 + 1)),

    IsPeriod = (v) =>
        if v is number then v >= 2000 and v <= 2100 and Number.Round(v) = v
        else if v is text then
            (let s = Text.Upper(Txt(v)), d = Text.Select(s, {"0".."9"}) in
                (Text.Length(d) = 4 and Text.StartsWith(d, "20"))
                or (Text.StartsWith(s, "HY") and Text.Length(d) = 2)
                or Text.StartsWith(s, "BENCHMARK"))
        else false,

    IsHeader = (r as number) as logical =>
        let cells = List.Transform(WinCols, (c) => Rows{r}{c}),
            nPer  = List.Count(List.Select(cells, IsPeriod)),
            nNum  = List.Count(List.Select(cells, (v) => v is number and not IsPeriod(v)))
        in nPer >= 2 and nNum = 0,
    HdrAll  = List.Buffer(List.Select(List.Numbers(0, NR), IsHeader)),
    NoNum    = (a, b) => List.AllTrue(List.Transform(List.Numbers(a + 1, List.Max({0, b - a - 1})), (r) =>
                   not List.AnyTrue(List.Transform(WinCols, (c) => Rows{r}{c} is number)))),
    IsSubHdr = (r) => List.AnyTrue(List.Transform(HdrAll, (p) => p < r and r - p <= 4 and NoNum(p, r))),
    HdrRows  = List.Buffer(List.Select(HdrAll, (r) => not IsSubHdr(r))),

    Junk = (t) => t = null or t = "" or t = "0" or Text.StartsWith(t, "*")
                  or Text.Upper(t) = "TO ALLOCATE" or Text.StartsWith(Text.Upper(t), "WHERE NEEDED"),
    FillFwd = (lst as list) as list => List.Accumulate(lst, {}, (s, v) =>
                  s & {if v <> null and v <> "" then v else (if List.IsEmpty(s) then null else List.Last(s))}),
    Prefixes = if SectionPrefix = null or SectionPrefix = "" then {} else Text.Split(SectionPrefix, "|"),
    IsSec = (t) => t <> null and List.AnyTrue(List.Transform(Prefixes, (p) => Text.StartsWith(t, p))),

    P0Of  = (h) => List.Min(List.Select(WinCols, (c) => IsPeriod(Rows{h}{c}))),
    EntOf = (h) =>
        let p0 = P0Of(h),
            LeftCols = List.Reverse(List.Select(WinCols, (c) => c < p0)),
            cells = List.Combine(List.Transform(List.Select(List.Numbers(h, 7, -1), (r) => r >= 0),
                        (r) => List.Transform(LeftCols, (c) => Txt(Rows{r}{c}))))
        in List.First(List.Select(cells, (t) => t <> null and
               (Text.StartsWith(Text.Upper(t), "COMPANY #")
                or List.AnyTrue(List.Transform(EntNames, (n) => Text.StartsWith(Text.Upper(t), n))))), null),
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

            SubR     = List.First(List.Select(HdrAll, (q) => q > h and q <= h + 4 and not List.Contains(HdrRows, q)), h + 1),
            SubOK    = SubR < hEnd,
            SubCells = if SubOK then List.Transform(ValArea, (c) => Rows{SubR}{c})
                       else List.Repeat({null}, List.Count(ValArea)),
            IsSubTxt = (v) => v is text
                       and (not IsPeriod(v) or Text.Length(Txt(v)) > 12)
                       and not List.Contains({"", "N.A.", "N.A", "-", "N/A"}, Text.Upper(Txt(v))),
            SubLab   = if SubOK and LabelCol <> null then Txt(Rows{SubR}{LabelCol}) else null,
            HasSub   = SubOK and List.Count(List.Select(SubCells, IsSubTxt)) >= 2
                       and (SubLab = null or SubLab = "" or Text.StartsWith(Text.Lower(SubLab), "amounts")),
            ValPos   = List.Select(List.Positions(ValArea), (i) =>
                           if HasSub then IsSubTxt(SubCells{i}) and HFill{i} <> null
                           else IsPeriod(HRow{ValArea{i}})),

            GCells  = if h > 0 then List.Transform(ValArea, (c) =>
                          let t = Txt(Rows{h - 1}{c}) in if t = null or t = "" or IsPeriod(t) then null else t)
                      else List.Repeat({null}, List.Count(ValArea)),
            UseGrp  = List.Count(List.RemoveNulls(GCells)) >= 2,
            GFill   = if UseGrp then FillFwd(GCells) else List.Repeat({null}, List.Count(ValArea)),

            BodyStart = if HasSub then SubR + 1 else h + 1,
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
            else if Text.StartsWith(s, "BENCHMARK") then "Benchmark"
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
