
Scale Factor = SELECTEDVALUE ( Scale[Multiplier], 1 )

Value =
SUMX (
    VALUES ( FactData[NoScale] ),
    VAR v = CALCULATE ( SUM ( FactData[Value] ) )
    RETURN IF ( FactData[NoScale], v, v * [Scale Factor] )
)

Ref Scenario Sel = SELECTEDVALUE ( 'Ref Scenario'[Scenario], "Current" )

Cmp Scenario Sel = SELECTEDVALUE ( 'Cmp Scenario'[Scenario], "Previous" )

Value Ref =
VAR s = [Ref Scenario Sel]
RETURN CALCULATE ( [Value], FactData[Scenario] = s )

Value Cmp =
VAR s = [Cmp Scenario Sel]
RETURN CALCULATE ( [Value], FactData[Scenario] = s )

Delta = [Value Ref] - [Value Cmp]

Delta % = DIVIDE ( [Delta], ABS ( [Value Cmp] ) )

Scale Label = "Amounts in EUR " & LOWER ( SELECTEDVALUE ( Scale[Scale], "Millions" ) )

Scenario Label = [Ref Scenario Sel] & " vs " & [Cmp Scenario Sel]


.............

Total Remittance =
CALCULATE ( [Value Ref], FactData[Topic] = "Remittance", FactData[IsTotal] = FALSE () )

Solvency Ratio =
VAR AC = CALCULATE ( [Value Ref], FactData[Topic] = "Solvency",
             FactData[Item] = "Available Capital (after Dividend if any)" )
VAR RC = CALCULATE ( [Value Ref], FactData[Topic] = "Solvency", FactData[Item] = "Required Capital" )
RETURN DIVIDE ( AC, RC )

Distribution Ratio =
VAR D = CALCULATE ( [Value Ref], FactData[Topic] = "Solvency", FactData[Item] = "Total Distributions (Accrual Basis)" )
VAR N = CALCULATE ( [Value Ref], FactData[Topic] = "Solvency", FactData[Item] = "Local GAAP Net Result" )
RETURN DIVIDE ( D, N )

............

IF ( SELECTEDVALUE ( FactData[NoScale], FALSE () ), "0.0%", "#,##0.0;(#,##0.0);-" )


.............



