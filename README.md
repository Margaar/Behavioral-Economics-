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
                   {"Measure", type text}}),
    Renamed  = Table.ReplaceValue(Typed, "amounts in Million of EUR", "Expected SCR Absorption",
                   Replacer.ReplaceValue, {"Block"}),
    NoDelta  = Table.SelectRows(Renamed, each
                   not Text.StartsWith([Block], "Delta Current")
                   and not Text.StartsWith([Item], "Delta Current"))
in
    NoDelta
