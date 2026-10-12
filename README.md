
let
    FromFacts = Table.Distinct(Table.SelectColumns(
                    Table.SelectRows(FactData, each [Entity] <> null and [Entity] <> ""), {"Entity"})),
    Joined    = Table.NestedJoin(FromFacts, {"Entity"}, EntityList, {"Entity"}, "L",
                    JoinKind.LeftOuter, Comparer.OrdinalIgnoreCase),
    Expanded  = Table.ExpandTableColumn(Joined, "L", {"Code", "Country"}),
    Flagged   = Table.AddColumn(Expanded, "InList", each [Code] <> null or [Country] <> null, type logical)
in
    Flagged
