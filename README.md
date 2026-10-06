=====================================================================
 FORMULA AUDIT -> POWER BI   |   Power Query code
 Reads the CSV files written by the RunAudit macro (AuditDump_VBA.txt)
=====================================================================

SETUP
  1. Gestisci parametri > Nuovo parametro
        Nome: AuditPath    Tipo: Testo
        Valore: the full path of the "AuditOut" folder (next to AuditTool.xlsm)
  2. Create the queries below, in this order
     (Nuova origine > Query vuota > Editor avanzato > paste > rename).
  3. Do NOT create relationships between these tables and FactData.

QUERY LIST
  fnAuditCsv           function          -
  AuditCells           all dumped cells  load OFF
  AuditStructure       sheets/names      load OFF
  AuditFindings        errors + flags in the NEW file          load ON
  FormulaDiff          cell-by-cell formula changes            load ON
  FormulaPatternDiff   formula changes that ignore position    load ON
  StructureDiff        sheets / names / links added, removed   load ON
  AuditSheets          one row per sheet of the NEW file       load ON


=====================================================================
 1) fnAuditCsv
=====================================================================

(prefix as text) as table =>
let
    Files  = Table.SelectRows(Folder.Files(AuditPath), each
                 Text.StartsWith([Name], prefix) and Text.Lower([Extension]) = ".csv"),
    Tables = List.Transform(Files[Content], each
                 Table.PromoteHeaders(
                     Csv.Document(_, [Delimiter = ",", Encoding = 1252, QuoteStyle = QuoteStyle.Csv]),
                     [PromoteAllScalars = true])),
    Result = if List.IsEmpty(Tables)
             then error "No audit files found in AuditPath. Run the RunAudit macro first."
             else Table.Combine(Tables)
in
    Result


=====================================================================
 2) AuditCells            (load OFF)
=====================================================================

let
    Source = fnAuditCsv("cells__"),
    Typed  = Table.TransformColumnTypes(Source, {{"Row", Int64.Type}, {"Col", Int64.Type}})
in
    Typed


=====================================================================
 3) AuditStructure        (load OFF)
=====================================================================

let
    Source = fnAuditCsv("structure__"),
    Typed  = Table.TransformColumnTypes(Source,
                 {{"Formulas", Int64.Type}, {"Errors", Int64.Type}, {"Flags", Int64.Type}})
in
    Typed


=====================================================================
 4) AuditFindings         (load ON)
    Every #REF!, #N/A, #VALUE!... and every suspicious cell of the NEW file.
    Needs no old file.
=====================================================================

let
    Rows = Table.SelectRows(AuditCells, each [Role] = "New" and ([Error] <> "" or [Flag] <> "")),
    Add  = Table.AddColumn(Rows, "Check", each
               if [Error] <> "" and [Flag] <> "" then "Error value " & [Error] & "; " & [Flag]
               else if [Error] <> "" then "Error value " & [Error]
               else [Flag], type text),
    Out  = Table.SelectColumns(Add,
               {"Run", "File", "Sheet", "Address", "Row", "Col", "Check", "Error", "Flag", "Formula"})
in
    Out


=====================================================================
 5) FormulaDiff           (load ON)
    Same sheet + same cell address, old file vs new file.
    Exact, but noisy when rows or columns were inserted.
=====================================================================

let
    Fx     = Table.SelectRows(AuditCells, each [Formula] <> ""),
    Old    = Table.RenameColumns(
                 Table.SelectColumns(Table.SelectRows(Fx, each [Role] = "Old"),
                     {"Run", "Sheet", "Address", "Formula"}),
                 {{"Run", "o.Run"}, {"Sheet", "o.Sheet"}, {"Address", "o.Address"}, {"Formula", "OldFormula"}}),
    RunsWithOld = List.Buffer(List.Distinct(Old[o.Run])),
    New    = Table.SelectColumns(
                 Table.SelectRows(Fx, each [Role] = "New" and List.Contains(RunsWithOld, [Run])),
                 {"Run", "Sheet", "Address", "Formula"}),
    Joined = Table.Join(New, {"Run", "Sheet", "Address"}, Old, {"o.Run", "o.Sheet", "o.Address"},
                 JoinKind.FullOuter),
    Status = Table.AddColumn(Joined, "Status", each
                 if [Formula] = null then "Formula removed or moved"
                 else if [OldFormula] = null then "Formula added"
                 else if [Formula] <> [OldFormula] then "Formula changed"
                 else "Same", type text),
    Diff   = Table.SelectRows(Status, each [Status] <> "Same"),
    A1     = Table.AddColumn(Diff, "RunX",     each if [Run] = null then [o.Run] else [Run], type text),
    A2     = Table.AddColumn(A1,   "SheetX",   each if [Sheet] = null then [o.Sheet] else [Sheet], type text),
    A3     = Table.AddColumn(A2,   "AddressX", each if [Address] = null then [o.Address] else [Address], type text),
    Sel    = Table.SelectColumns(A3, {"RunX", "SheetX", "AddressX", "Status", "Formula", "OldFormula"}),
    Out    = Table.RenameColumns(Sel,
                 {{"RunX", "Run"}, {"SheetX", "Sheet"}, {"AddressX", "Address"}, {"Formula", "NewFormula"}})
in
    Out


=====================================================================
 6) FormulaPatternDiff    (load ON)
    Counts how many cells of each sheet use each formula, old vs new.
    A copied formula has the same R1C1 text wherever it sits, so this
    view ignores position: "a formula that did not exist before", or
    "one cell fewer uses this formula than last year".
=====================================================================

let
    Fx      = Table.SelectRows(AuditCells, each [Formula] <> ""),
    RunsWithOld = List.Buffer(List.Distinct(Table.SelectRows(Fx, each [Role] = "Old")[Run])),
    InScope = Table.SelectRows(Fx, each List.Contains(RunsWithOld, [Run])),
    Grouped = Table.Group(InScope, {"Run", "Sheet", "Formula"}, {
                  {"CellsNew", each List.Count(List.Select([Role], (r) => r = "New")), Int64.Type},
                  {"CellsOld", each List.Count(List.Select([Role], (r) => r = "Old")), Int64.Type},
                  {"ExampleCell", each List.First([Address]), type text}}),
    Changed = Table.SelectRows(Grouped, each [CellsNew] <> [CellsOld]),
    Status  = Table.AddColumn(Changed, "Status", each
                  if [CellsOld] = 0 then "New formula (not in old file)"
                  else if [CellsNew] = 0 then "Formula no longer used"
                  else if [CellsNew] < [CellsOld] then "Used in fewer cells"
                  else "Used in more cells", type text),
    Delta   = Table.AddColumn(Status, "CellsDelta", each [CellsNew] - [CellsOld], Int64.Type)
in
    Delta


=====================================================================
 7) StructureDiff         (load ON)
    Sheets, named ranges and external links that were added, removed
    or changed (used range, visibility, number of formulas, definition).
=====================================================================

let
    S       = AuditStructure,
    RunsWithOld = List.Buffer(List.Distinct(Table.SelectRows(S, each [Role] = "Old")[Run])),
    InScope = Table.SelectRows(S, each List.Contains(RunsWithOld, [Run])),
    Grouped = Table.Group(InScope, {"Run", "Kind", "Name"}, {
                  {"InNew", each List.Contains([Role], "New"), type logical},
                  {"InOld", each List.Contains([Role], "Old"), type logical},
                  {"InfoNew", each List.First(Table.SelectRows(_, (x) => x[Role] = "New")[Info], null), type nullable text},
                  {"InfoOld", each List.First(Table.SelectRows(_, (x) => x[Role] = "Old")[Info], null), type nullable text},
                  {"FormulasNew", each List.Sum(Table.SelectRows(_, (x) => x[Role] = "New")[Formulas]), Int64.Type},
                  {"FormulasOld", each List.Sum(Table.SelectRows(_, (x) => x[Role] = "Old")[Formulas]), Int64.Type}}),
    Status  = Table.AddColumn(Grouped, "Status", each
                  if not [InOld] then "Added"
                  else if not [InNew] then "Removed"
                  else if [InfoNew] <> [InfoOld] or [FormulasNew] <> [FormulasOld] then "Changed"
                  else "Same", type text),
    Out     = Table.SelectRows(Status, each [Status] <> "Same")
in
    Out


=====================================================================
 8) AuditSheets           (load ON)
    One row per sheet of the NEW file: formulas, errors, flags.
=====================================================================

let
    Rows = Table.SelectRows(AuditStructure, each [Role] = "New" and [Kind] = "Sheet"),
    Out  = Table.SelectColumns(Rows, {"Run", "File", "Name", "Info", "Formulas", "Errors", "Flags"})
in
    Out


=====================================================================
 REPORT PAGE "Formula Audit"
=====================================================================

  Slicers : AuditFindings[Run] (one per audited file), [Sheet]
  Cards   : count of rows in AuditFindings, FormulaPatternDiff, StructureDiff
  Table 1 : AuditFindings        Sheet, Address, Check, Formula
  Table 2 : FormulaPatternDiff   Sheet, Status, Formula, CellsOld, CellsNew, ExampleCell
  Table 3 : StructureDiff        Kind, Name, Status, InfoOld, InfoNew
  Table 4 : FormulaDiff          Sheet, Address, Status, OldFormula, NewFormula
            (use it to look up a cell, not to read top to bottom)
  Bar     : AuditSheets          Name by Errors and Flags
