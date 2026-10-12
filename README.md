let
    Sheet = Workbook{[Item = "SAA", Kind = "Sheet"]}[Data],
    T     = fnReadSheet(Sheet, "Cash & Cash Equivalent", null, false, null),
    G     = Table.Group(T, {"BlockNo", "Block", "Entity"}, {
                {"Rows",     each Table.RowCount(_), Int64.Type},
                {"FromRow",  each List.Min([RowNo]), Int64.Type},
                {"ToRow",    each List.Max([RowNo]), Int64.Type},
                {"Periods",  each Text.Combine(List.Distinct(List.RemoveNulls([Period])), " | "), type text},
                {"Measures", each Text.Combine(List.Distinct(List.RemoveNulls([Measure])), " | "), type text},
                {"Items",    each Text.Combine(List.FirstN(List.Distinct([Item]), 3), " | "), type text}})
in
    G
