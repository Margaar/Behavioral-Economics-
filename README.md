=====================================================================
 SOLVENCY AoM — FIX
 Строки с ">>" 
=====================================================================

>> ШАГ 1. ДВЕ НОВЫЕ КОЛОНКИ в FactData
>> (вид Dati > FactData > Nuova colonna), по одной.

>> AoM Category: одна категория для EOF и SCR, как в таблице
>> "Solvency Ratio" в Excel (Tax impact EOF + Tax impact SCR = одна строка).

AoM Category =
IF ( FactData[Topic] = "Solvency AoM",
    VAR it = FactData[Item]
    RETURN
        SWITCH ( TRUE (),
            CONTAINSSTRING ( it, "(BoP)" ), "Solvency Ratio (BoP)",
            CONTAINSSTRING ( it, "(EoP)" ), "Solvency Ratio (EoP)",
            CONTAINSSTRING ( it, "Normalised" ), "Normalised Capital Generation",
            it ) )

>> AoM Category Order: порядок строк как в Excel.

AoM Category Order =
IF ( FactData[Topic] = "Solvency AoM",
    VAR it = FactData[Item]
    RETURN
        SWITCH ( TRUE (),
            CONTAINSSTRING ( it, "(BoP)" ), 0,
            CONTAINSSTRING ( it, "Regulatory" ), 1,
            CONTAINSSTRING ( it, "SII Review" ), 2,
            CONTAINSSTRING ( it, "Normalised" ), 3,
            CONTAINSSTRING ( it, "SII VNP" ), 4,
            CONTAINSSTRING ( it, "CY Tech" ), 5,
            CONTAINSSTRING ( it, "Non-Economic" ), 8,
            CONTAINSSTRING ( it, "Economic variances" ), 6,
            CONTAINSSTRING ( it, "YE25 to HY26" ) && NOT CONTAINSSTRING ( it, "SAA" ), 7,
            CONTAINSSTRING ( it, "SAA" ) && CONTAINSSTRING ( it, "consumed" ), 9,
            CONTAINSSTRING ( it, "SAA" ), 10,
            CONTAINSSTRING ( it, "Residual" ), 11,
            CONTAINSSTRING ( it, "Tax" ), 12,
            CONTAINSSTRING ( it, "M&A" ), 13,
            CONTAINSSTRING ( it, "Capital movements" ), 14,
            50 ) )

>> Потом: выделить колонку AoM Category > Ordina per colonna > AoM Category Order.


>> ШАГ 2. ЗАМЕНИТЬ 4 МЕРЫ BoP/EoP И 2 ПРОВЕРКИ
>> (нажать на меру, выделить всю формулу, вставить новую).
>> Добавлено: снятие фильтра с новых колонок и исключение строк "o/w".

AoM EOF BoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "EOF", CONTAINSSTRING ( FactData[Item], "(BoP)" ) )

AoM EOF EoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "EOF", CONTAINSSTRING ( FactData[Item], "(EoP)" ) )

AoM SCR BoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "SCR", CONTAINSSTRING ( FactData[Item], "(BoP)" ) )

AoM SCR EoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "SCR", CONTAINSSTRING ( FactData[Item], "(EoP)" ) )

AoM Check EOF =
[AoM EOF BoP]
+ CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
      REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
      FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "EOF",
      NOT CONTAINSSTRING ( FactData[Item], "(BoP)" ), NOT CONTAINSSTRING ( FactData[Item], "(EoP)" ),
      NOT CONTAINSSTRING ( FactData[Item], "o/w" ) )
- [AoM EOF EoP]

AoM Check SCR =
[AoM SCR BoP]
+ CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
      REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
      FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "SCR",
      NOT CONTAINSSTRING ( FactData[Item], "(BoP)" ), NOT CONTAINSSTRING ( FactData[Item], "(EoP)" ),
      NOT CONTAINSSTRING ( FactData[Item], "o/w" ) )
- [AoM SCR EoP]


>> ШАГ 3. ЗАМЕНИТЬ AoM Ratio Walk (строки "o/w" теперь не считаются,
>> иначе водопад считал бы их дважды)

AoM Ratio Walk =
VAR ln  = SELECTEDVALUE ( FactData[AoM Line] )
VAR e0  = [AoM EOF BoP]
VAR e1  = [AoM EOF EoP]
VAR s0  = [AoM SCR BoP]
VAR s1  = [AoM SCR EoP]
VAR v   = [Value Current]
VAR isBoP = CONTAINSSTRING ( ln, "(BoP)" )
VAR isEoP = CONTAINSSTRING ( ln, "(EoP)" )
VAR isSub = CONTAINSSTRING ( ln, "| o/w" )
RETURN
    SWITCH ( TRUE (),
        NOT HASONEVALUE ( FactData[AoM Line] ), BLANK (),
        isEoP || isSub, BLANK (),
        isBoP && LEFT ( ln, 3 ) = "EOF", DIVIDE ( e0, s0 ) * 100,
        isBoP, BLANK (),
        LEFT ( ln, 3 ) = "EOF", DIVIDE ( v, s0 ) * 100,
        LEFT ( ln, 3 ) = "SCR", DIVIDE ( v, s1 - s0 ) * ( DIVIDE ( e1, s1 ) - DIVIDE ( e1, s0 ) ) * 100,
        BLANK () )


>> ШАГ 4. НОВАЯ МЕРА: вклад категории в Solvency Ratio (как в Excel)
>> Strumenti misura > Formato: Percentuale, 1 десятичный знак.

AoM Ratio by Category =
VAR cat = SELECTEDVALUE ( FactData[AoM Category] )
VAR pp =
    IF ( cat = "Solvency Ratio (EoP)", [AoM Ratio EoP] * 100,
         SUMX ( VALUES ( FactData[AoM Line] ), [AoM Ratio Walk] ) )
RETURN IF ( HASONEVALUE ( FactData[AoM Category] ), DIVIDE ( pp, 100 ) )


>> ШАГ 5. НОВЫЙ AoM Comment (заменить целиком).
>> Без выбора года берёт последний год; клик по ячейке/срез Year = этот год.

AoM Comment =
VAR hardLim = 1.10
VAR softLim = 1.30
VAR fmt  = "#,##0"
VAR fmtP = "+0.0;-0.0;0.0"
VAR nl   = UNICHAR ( 10 )
VAR y    = MAX ( FactData[Year] )
VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency AoM" )
VAR who  = IF ( nEnt = 1,
               CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency AoM" ),
               nEnt & " companies (simple sum)" )
VAR r0   = CALCULATE ( [AoM Ratio BoP], FactData[Year] = y )
VAR r1   = CALCULATE ( [AoM Ratio EoP], FactData[Year] = y )
VAR rSol = CALCULATE ( [SOL Ratio], FactData[Year] = y )
VAR st =
    IF ( ISBLANK ( r1 ), "cannot be calculated",
    IF ( r1 < hardLim, "BELOW the " & FORMAT ( hardLim, "0%" ) & " threshold: action needed",
    IF ( r1 < softLim, "between the " & FORMAT ( hardLim, "0%" ) & " and " & FORMAT ( softLim, "0%" ) & " thresholds",
         "above the " & FORMAT ( softLim, "0%" ) & " threshold" ) ) )
VAR lines =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( FactData[AoM Line] ), "@pp", [AoM Ratio Walk] ),
        REMOVEFILTERS ( DimItem ),
        REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
        FactData[Topic] = "Solvency AoM", FactData[Year] = y )
VAR eofPP = SUMX ( FILTER ( lines, LEFT ( FactData[AoM Line], 3 ) = "EOF" && NOT CONTAINSSTRING ( FactData[AoM Line], "(BoP)" ) ), [@pp] )
VAR scrPP = SUMX ( FILTER ( lines, LEFT ( FactData[AoM Line], 3 ) = "SCR" ), [@pp] )
VAR cats =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( FactData[AoM Category] ), "@pp", [AoM Ratio by Category] * 100 ),
        REMOVEFILTERS ( DimItem ),
        REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order], FactData[AoM Category], FactData[AoM Category Order] ),
        FactData[Topic] = "Solvency AoM", FactData[Year] = y )
VAR movers  = FILTER ( cats, NOT CONTAINSSTRING ( FactData[AoM Category], "Solvency Ratio" ) && [@pp] <> 0 )
VAR top3Txt = CONCATENATEX ( TOPN ( 3, movers, ABS ( [@pp] ), DESC ),
                  FactData[AoM Category] & " (" & FORMAT ( [@pp], fmtP ) & " pp)", "; ", ABS ( [@pp] ), DESC )
VAR chkE = ROUND ( CALCULATE ( [AoM Check EOF], FactData[Year] = y ), 0 )
VAR chkS = ROUND ( CALCULATE ( [AoM Check SCR], FactData[Year] = y ), 0 )
VAR pageText =
    "SOLVENCY RATIO MOVEMENT " & y & ": " & UPPER ( who ) & nl & nl
    & "Solvency ratio moved from " & FORMAT ( r0, "0.0%" ) & " to " & FORMAT ( r1, "0.0%" )
    & " (" & FORMAT ( ( r1 - r0 ) * 100, fmtP ) & " pp), " & st & "." & nl
    & "Own funds movements: " & FORMAT ( eofPP, fmtP ) & " pp; SCR movements: " & FORMAT ( scrPP, fmtP ) & " pp." & nl
    & IF ( top3Txt <> "", "Main drivers: " & top3Txt & ".", "No movements reported." )
    & nl & nl
    & IF ( chkE = 0 && chkS = 0, "Check: EOF and SCR movements reconcile with the closing balances.",
           "Check: movements do not reconcile (EOF diff " & FORMAT ( chkE, fmt ) & ", SCR diff " & FORMAT ( chkS, fmt ) & "). Please review the input file." )
    & IF ( NOT ISBLANK ( rSol ) && ROUND ( ( r1 - rSol ) * 1000, 0 ) <> 0,
           nl & "Note: closing ratio differs from the Solvency sheet (" & FORMAT ( rSol, "0.0%" ) & ").", "" )
VAR cat = SELECTEDVALUE ( FactData[AoM Category] )
VAR catLines =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( FactData[AoM Line] ), "@pp", [AoM Ratio Walk], "@v", [Value Current] ),
        FactData[Topic] = "Solvency AoM", FactData[Year] = y )
VAR eofV = SUMX ( FILTER ( catLines, LEFT ( FactData[AoM Line], 3 ) = "EOF" ), [@v] )
VAR scrV = SUMX ( FILTER ( catLines, LEFT ( FactData[AoM Line], 3 ) = "SCR" ), [@v] )
VAR eofP = SUMX ( FILTER ( catLines, LEFT ( FactData[AoM Line], 3 ) = "EOF" ), [@pp] )
VAR scrP = SUMX ( FILTER ( catLines, LEFT ( FactData[AoM Line], 3 ) = "SCR" ), [@pp] )
VAR rowText =
    UPPER ( cat ) & " " & y & ": " & who & nl & nl
    & IF ( CONTAINSSTRING ( cat, "Solvency Ratio" ),
           "Ratio: " & FORMAT ( IF ( CONTAINSSTRING ( cat, "(BoP)" ), r0, r1 ), "0.0%" ) & ".",
           "Impact on the solvency ratio: " & FORMAT ( eofP + scrP, fmtP ) & " pp." & nl
           & "Own funds: " & FORMAT ( eofV + 0, fmt ) & " (" & FORMAT ( eofP + 0, fmtP ) & " pp); "
           & "SCR: " & FORMAT ( scrV + 0, fmt ) & " (" & FORMAT ( scrP + 0, fmtP ) & " pp)."
           & IF ( CONTAINSSTRING ( cat, "o/w" ), nl & "Sub-item: already included in the line above.", "" ) )
RETURN IF ( HASONEVALUE ( FactData[AoM Category] ), rowText, pageText )
