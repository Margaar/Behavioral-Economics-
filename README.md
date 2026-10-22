let
    F  = Table.SelectRows(Files, each Text.Contains([Name], "Capital Deep Dive")),
    Wb = Excel.Workbook(F{0}[Content], null, true)
in
    Wb
