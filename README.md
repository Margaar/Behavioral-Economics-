let
    Src  = Table.SelectRows(
               Table.SelectColumns(FactData, {"LineKey", "Topic", "Block", "Item", "RowNo"}),
               each [LineKey] <> null and [LineKey] <> ""),
    Uniq = Table.Group(Src, {"LineKey"},
               {{"Topic",     each List.First([Topic]), type text},
                {"Block",     each List.First([Block]), type text},
                {"Item",      each List.First([Item]),  type text},
                {"LineOrder", each List.Min([RowNo]),   Int64.Type}},
               GroupKind.Global, Comparer.OrdinalIgnoreCase),
    ByIt = Table.Group(Uniq, {"Item"},
               {{"T", each Table.AddIndexColumn(Table.Sort(_, {{"LineOrder", Order.Ascending}}),
                                                "Dup", 0, 1, Int64.Type)}},
               GroupKind.Global, Comparer.OrdinalIgnoreCase),
    Exp  = Table.ExpandTableColumn(ByIt, "T", {"LineKey", "Topic", "Block", "LineOrder", "Dup"}),
    Line = Table.AddColumn(Exp, "Line", each
               (if [Item] = null then "" else [Item]) & Text.Repeat(Character.FromNumber(8203), [Dup]), type text),
    BOrd = Table.Group(Line, {"Block"}, {{"BlockOrder", each List.Min([LineOrder]), Int64.Type}}),
    Join = Table.NestedJoin(Line, {"Block"}, BOrd, {"Block"}, "B", JoinKind.LeftOuter),
    Out  = Table.ExpandTableColumn(Join, "B", {"BlockOrder"})
in
    Out
