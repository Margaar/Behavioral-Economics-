    LK     = Table.Group(WithKey, {"LineKey", "Item", "Block"},
                 {{"LineOrder", each List.Min([RowNo]), Int64.Type}}),
    ByIt   = Table.Group(LK, {"Item"},
                 {{"T", each Table.AddIndexColumn(Table.Sort(_, {{"LineOrder", Order.Ascending}}),
                                                  "Dup", 0, 1, Int64.Type)}}),
    LK2    = Table.ExpandTableColumn(ByIt, "T", {"LineKey", "Block", "LineOrder", "Dup"}),
    LK3    = Table.AddColumn(LK2, "Line", each
                 (if [Item] = null then "" else [Item]) & Text.Repeat(Character.FromNumber(8203), [Dup]), type text),
    BOrd   = Table.Group(LK3, {"Block"}, {{"BlockOrder", each List.Min([LineOrder]), Int64.Type}}),
    LK4    = Table.NestedJoin(LK3, {"Block"}, BOrd, {"Block"}, "B", JoinKind.LeftOuter),
    LK5    = Table.SelectColumns(Table.ExpandTableColumn(LK4, "B", {"BlockOrder"}),
                 {"LineKey", "Line", "LineOrder", "BlockOrder"}),
    J      = Table.NestedJoin(WithKey, {"LineKey"}, LK5, {"LineKey"}, "L", JoinKind.LeftOuter),
    Final  = Table.ExpandTableColumn(J, "L", {"Line", "LineOrder", "BlockOrder"})
in
    Final
