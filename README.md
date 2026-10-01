beh ec


let
    Files    = Table.SelectRows(Folder.Files(FolderPath), each
                   (Text.Lower([Extension]) = ".xlsm" or Text.Lower([Extension]) = ".xlsx")
                   and not Text.StartsWith([Name], "~$")),
    WithWb   = Table.AddColumn(Files, "Wb", each Excel.Workbook([Content], null, true)),
    Keep     = Table.SelectColumns(WithWb, {"Name", "Wb"}),
    WithCfg  = Table.AddColumn(Keep, "Cfg", each BlockConfig),
    ExpCfg   = Table.ExpandTableColumn(WithCfg, "Cfg",
                   {"Topic", "Sheet", "FromText", "ToText", "SideGroup", "SectionPrefix"}),
    WithData = Table.AddColumn(ExpCfg, "Data", each
                   try fnReadSheet([Wb]{[Item = [Sheet], Kind = "Sheet"]}[Data],
                                   [FromText], [ToText], [SideGroup], [SectionPrefix])
                   otherwise null),
    Found    = Table.SelectRows(WithData, each [Data] <> null),
    Expanded = Table.ExpandTableColumn(Found, "Data",
                   {"BlockNo", "Block", "Entity", "Group", "Scenario", "Section", "RowNo",
                    "Item", "SignHint", "IsTotal", "Period", "PeriodType", "Year", "Measure", "Value"}),
    AddBU    = Table.AddColumn(Expanded, "BU", each Text.BeforeDelimiter([Name], "."), type text),
    Removed  = Table.RemoveColumns(AddBU, {"Wb", "FromText", "ToText", "SideGroup", "SectionPrefix"}),
    Typed    = Table.TransformColumnTypes(Removed, {
                   {"BlockNo", Int64.Type}, {"RowNo", Int64.Type}, {"Year", Int64.Type},
                   {"Value", type number}, {"IsTotal", type logical}})
in
    Typed
