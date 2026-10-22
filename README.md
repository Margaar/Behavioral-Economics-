You are helping me map accounts from a central reporting extract to the lines of a capital report.

Sheet "CMP_Accounts" lists accounts from the central system: Process, AccountCode, Description, Rows, Sum EUR m.
Sheet "CDD_Items" lists the report lines: Topic, Item. Each Item is the exact row label used in the report.

Create a new sheet "AccountMap" with exactly these columns:
AccountCode | Topic | Item | Sign | Order | IsTotal | Confidence | Reason

Rules:
1. One row per AccountCode from CMP_Accounts (ignore rows with an empty AccountCode).
2. Topic and Item must be copied EXACTLY (same spelling, spaces, brackets) from a row of "CDD_Items". Never invent a new Item.
3. Map an account only when its Description clearly means the same thing as the Item, e.g.
   "Dividends Outflow (excl. interim dividends) (-) cash view basis" -> "Dividends Outflow (excl. interim dividends)";
   "Interim Dividends Outflow (-) cash view basis" -> the interim dividends line;
   "Local GAAP Net Result" -> "Local GAAP Net Result";
   Eligible Own Funds total -> the "Eligible Own Funds" line of Topic "Solvency";
   Solvency Capital Requirement total -> the "Solvency Capital Requirement" line of Topic "Solvency".
   Prefer totals over sub-components; map sub-components only if the report has that exact sub-line.
4. If several accounts could match one Item, keep the best one and set the others to Topic = "REVIEW".
5. If no clear match exists, set Topic = "REVIEW" and leave Item empty.
6. Sign = 1 when the account and the report line use the same sign convention (both outflows negative); Sign = -1 only if the Description shows the opposite sign to the report line. When unsure, use 1.
7. Order = the position of the Item within its Topic (1, 2, 3...) in the order of "CDD_Items".
8. IsTotal = TRUE only if the Item starts with "Total"; otherwise FALSE.
9. Confidence = High / Medium / Low; Reason = one short sentence explaining the match.
Output the sheet as an Excel table named AccountMap.





let
    F   = Table.SelectRows(Files, each Text.Contains([Name], "AccountMap")),
    Wb  = Excel.Workbook(F{0}[Content], null, true),
    T   = Wb{[Item = "AccountMap", Kind = "Table"]}[Data],
    Sel = Table.SelectColumns(T, {"AccountCode", "Topic", "Item", "Sign", "Order", "IsTotal"}),
    Ok  = Table.SelectRows(Sel, each [Topic] <> null and [Topic] <> "REVIEW"),
    Typ = Table.TransformColumnTypes(Ok, {{"Sign", Int64.Type}, {"Order", Int64.Type}, {"IsTotal", type logical}})
in
    Typ
