let
    FromFacts = Table.Distinct(Table.SelectColumns(
                    Table.SelectRows(FactData, each [Entity] <> null and [Entity] <> ""), {"Entity"})),
    FactsKey  = Table.AddColumn(FromFacts, "Key", each Text.Upper(Text.Trim([Entity])), type text),
    ListKey   = Table.AddColumn(EntityList, "Key", each Text.Upper(Text.Trim([Entity])), type text),
    ListSlim  = Table.SelectColumns(ListKey, {"Key", "Code", "Country"}),
    Joined    = Table.NestedJoin(FactsKey, {"Key"}, ListSlim, {"Key"}, "L", JoinKind.LeftOuter),
    Expanded  = Table.ExpandTableColumn(Joined, "L", {"Code", "Country"}),
    Flagged   = Table.AddColumn(Expanded, "InList", each [Code] <> null or [Country] <> null, type logical),
    Result    = Table.RemoveColumns(Flagged, {"Key"})
in
    Result
