=====================================================================
 SMART COMMENTS  (Solvency / Capital Needs / Total Remittance)
 Home > Nuova misura : one measure at a time.
 Every comment works the same way:
   - nothing clicked        -> summary of the whole page
   - click a row            -> comment on that line (by year, vs previous plan)
   - click a cell (row+year)-> comment on that single year
=====================================================================


---------------------------------------------------------------------
 A) SOLVENCY, DISTRIBUTIONS & FTC
 Helper measures first (8 of them), then SOL Comment.
---------------------------------------------------------------------

SOL EOF =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Eligible Own Funds" ) )

SOL SCR =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Solvency Capital Requirement" ) )

SOL Ratio = DIVIDE ( [SOL EOF], [SOL SCR] )

SOL Ratio Prev =
DIVIDE (
    CALCULATE ( [Value Previous], REMOVEFILTERS ( DimItem ),
        FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Eligible Own Funds" ) ),
    CALCULATE ( [Value Previous], REMOVEFILTERS ( DimItem ),
        FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Solvency Capital Requirement" ) ) )

SOL FTC =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Free Tangible Capital" ) )

SOL NetResult =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Net Result" ) )

SOL Distributions =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Total Distributions" ) )

SOL OtherDistr =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Other distributions" ) )

SOL ForeseeableDiv =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Foreseeable" ) )

SOL CapNeeds =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Solvency", CONTAINSSTRING ( FactData[Item], "Capital Needs" ) )


SOL Comment =
-- thresholds: take them from the Local RAF sheet of the company (change here if different)
VAR hardLim = 1.10
VAR softLim = 1.30
VAR fmt  = "#,##0"
VAR fmtD = "+#,##0;-#,##0;0"
VAR nl   = UNICHAR ( 10 )

VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( DimItem ),
              FactData[Topic] = "Solvency", NOT ISBLANK ( FactData[Value] ) )
VAR y0 = MINX ( yrs, FactData[Year] )
VAR y1 = MAXX ( yrs, FactData[Year] )
VAR period = IF ( y0 = y1, "in " & y1, y0 & "-" & y1 )

VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency" )
VAR who  = IF ( nEnt = 1,
               CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency" ),
               nEnt & " companies (simple sum, no diversification)" )

-- solvency ratio
VAR eof1 = CALCULATE ( [SOL EOF], FactData[Year] = y1 )
VAR scr1 = CALCULATE ( [SOL SCR], FactData[Year] = y1 )
VAR r1   = DIVIDE ( eof1, scr1 )
VAR r0   = CALCULATE ( [SOL Ratio], FactData[Year] = y0 )
VAR rP1  = CALCULATE ( [SOL Ratio Prev], FactData[Year] = y1 )
VAR ratioTbl = FILTER ( ADDCOLUMNS ( yrs, "@r", CALCULATE ( [SOL Ratio] ) ), NOT ISBLANK ( [@r] ) )
VAR minRow = TOPN ( 1, ratioTbl, [@r], ASC )
VAR minR   = MAXX ( minRow, [@r] )
VAR minY   = MAXX ( minRow, FactData[Year] )
VAR nBelow = COUNTROWS ( FILTER ( ratioTbl, [@r] < hardLim ) )
VAR status =
    SWITCH ( TRUE (),
        ISBLANK ( r1 ), "cannot be calculated (SCR missing)",
        r1 < hardLim, "BELOW the " & FORMAT ( hardLim, "0%" ) & " threshold: action needed",
        r1 < softLim, "between the " & FORMAT ( hardLim, "0%" ) & " and " & FORMAT ( softLim, "0%" ) & " thresholds: limited buffer",
        "above the " & FORMAT ( softLim, "0%" ) & " threshold: comfortable buffer" )

-- distributions, capital needs, FTC (flows = sum of the period, stocks = last year)
VAR nr   = [SOL NetResult]
VAR dist = COALESCE ( [SOL Distributions], 0 )
VAR oth  = COALESCE ( [SOL OtherDistr], 0 )
VAR fd   = COALESCE ( [SOL ForeseeableDiv], 0 )
VAR capn = COALESCE ( [SOL CapNeeds], 0 )
VAR payout = DIVIDE ( -dist, nr )
VAR ftc0 = CALCULATE ( [SOL FTC], FactData[Year] = y0 )
VAR ftc1 = CALCULATE ( [SOL FTC], FactData[Year] = y1 )

VAR pageText =
    "SOLVENCY, DISTRIBUTIONS & FTC: " & UPPER ( who ) & ", " & period & nl & nl
    & IF ( ISBLANK ( eof1 ) && ISBLANK ( scr1 ),
           "No solvency data for this selection.",
           "Solvency ratio " & y1 & ": " & FORMAT ( r1, "0.0%" )
           & " (EOF " & FORMAT ( eof1 + 0, fmt ) & " / SCR " & FORMAT ( scr1 + 0, fmt ) & "), " & status & "."
           & IF ( y0 <> y1 && NOT ISBLANK ( r0 ),
                  nl & "Change since " & y0 & ": " & FORMAT ( ( r1 - r0 ) * 100, "+0.0;-0.0;0.0" )
                  & " pp (from " & FORMAT ( r0, "0.0%" ) & ").", "" )
           & IF ( y0 <> y1 && NOT ISBLANK ( minR ),
                  nl & "Lowest ratio in the plan: " & FORMAT ( minR, "0.0%" ) & " in " & minY
                  & IF ( nBelow > 0, "; " & nBelow & " year(s) below " & FORMAT ( hardLim, "0%" ), "" ) & ".", "" )
           & IF ( NOT ISBLANK ( rP1 ),
                  nl & "Versus previous plan (" & y1 & "): " & FORMAT ( ( r1 - rP1 ) * 100, "+0.0;-0.0;0.0" )
                  & " pp (previous " & FORMAT ( rP1, "0.0%" ) & ").",
                  nl & "No previous plan available for comparison." ) )
    & nl & nl
    & "Local GAAP net result " & period & ": " & FORMAT ( nr + 0, fmt ) & "." & nl
    & "Distributions (accrual basis): " & FORMAT ( -dist, fmt )
    & IF ( nr > 0 && dist <> 0, " (" & FORMAT ( payout, "0%" ) & " of net result)", "" )
    & "; other distributions: " & FORMAT ( -oth, fmt )
    & "; foreseeable dividends: " & FORMAT ( -fd, fmt ) & "." & nl
    & IF ( capn <> 0, "Capital needs (inflows) " & period & ": " & FORMAT ( capn, fmt ) & ".",
                      "No capital needs planned." ) & nl
    & "Free Tangible Capital " & y1 & ": " & FORMAT ( ftc1 + 0, fmt )
    & IF ( y0 <> y1, " (" & FORMAT ( ftc1 - ftc0, fmtD ) & " vs " & y0 & ")", "" ) & "."

-- one line selected
VAR it  = SELECTEDVALUE ( FactData[Item] )
VAR lbl = COALESCE ( SELECTEDVALUE ( DimItem[Item Label] ), it )
VAR isStock = CONTAINSSTRING ( it, "Own Funds" ) || CONTAINSSTRING ( it, "Capital Requirement" )
              || CONTAINSSTRING ( it, "Tangible" )
VAR v1   = CALCULATE ( [Value Current], FactData[Year] = y1 )
VAR v0   = CALCULATE ( [Value Current], FactData[Year] = y0 )
VAR vSum = [Value Current]
VAR vP1  = CALCULATE ( [Value Previous], FactData[Year] = y1 )
VAR byYear = CONCATENATEX ( yrs, FactData[Year] & ": " & FORMAT ( CALCULATE ( [Value Current] ) + 0, fmt ),
                 " | ", FactData[Year], ASC )
VAR meaning =
    SWITCH ( TRUE (),
        CONTAINSSTRING ( it, "Own Funds" ), "Capital available to absorb losses under Solvency II.",
        CONTAINSSTRING ( it, "Capital Requirement" ), "Capital required by Solvency II to cover the company's risks.",
        CONTAINSSTRING ( it, "Tangible" ), "Tangible capital in excess of requirements, a source for future remittances.",
        CONTAINSSTRING ( it, "Net Result" ), "Statutory profit: the base for dividend capacity.",
        CONTAINSSTRING ( it, "Foreseeable" ), "Dividends already deducted from own funds.",
        CONTAINSSTRING ( it, "Distributions" ), "Capital returned to the shareholder.",
        CONTAINSSTRING ( it, "Capital Needs" ), "Capital injections requested from the shareholder.",
        "" )

VAR rowText =
    UPPER ( lbl ) & ": " & who & nl & nl
    & IF ( y0 = y1,
           "Value " & y1 & ": " & FORMAT ( v1 + 0, fmt )
           & IF ( vP1 <> 0, " vs previous plan " & FORMAT ( vP1, fmt ) & " (" & FORMAT ( v1 - vP1, fmtD ) & ")", "" ) & ".",
           "By year: " & byYear & "." & nl
           & IF ( isStock,
                  "Change " & y0 & "-" & y1 & ": " & FORMAT ( v1 - v0, fmtD )
                  & IF ( v0 <> 0, " (" & FORMAT ( DIVIDE ( v1 - v0, ABS ( v0 ) ), "+0%;-0%;0%" ) & ")", "" ) & ".",
                  "Total " & period & ": " & FORMAT ( vSum + 0, fmt ) & "." )
           & IF ( vP1 <> 0, nl & "Versus previous plan in " & y1 & ": " & FORMAT ( v1 - vP1, fmtD ) & ".", "" ) )
    & IF ( meaning <> "", nl & nl & meaning, "" )
    & IF ( vSum = 0, nl & "No amounts planned for this line.", "" )

RETURN IF ( HASONEVALUE ( FactData[Item] ), rowText, pageText )


---------------------------------------------------------------------
 B) CAPITAL NEEDS
 2 helper measures, then CN Comment.
---------------------------------------------------------------------

CN Total Cur =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ),
    FactData[Topic] = "Capital Needs", FactData[IsTotal] = FALSE () )

CN Total Prev =
CALCULATE ( [Value Previous], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ),
    FactData[Topic] = "Capital Needs", FactData[IsTotal] = FALSE () )


CN Comment =
VAR fmt  = "#,##0"
VAR fmtD = "+#,##0;-#,##0;0"
VAR nl   = UNICHAR ( 10 )

VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ),
              FactData[Topic] = "Capital Needs", NOT ISBLANK ( FactData[Value] ) )
VAR y0 = MINX ( yrs, FactData[Year] )
VAR y1 = MAXX ( yrs, FactData[Year] )
VAR period = IF ( y0 = y1, "in " & y1, y0 & "-" & y1 )

-- page level
VAR yearTbl = ADDCOLUMNS ( yrs, "@c", [CN Total Cur], "@p", [CN Total Prev] )
VAR tot     = SUMX ( yearTbl, [@c] )
VAR peakRow = TOPN ( 1, yearTbl, [@c], DESC )
VAR peakY   = MAXX ( peakRow, FactData[Year] )
VAR peakV   = MAXX ( peakRow, [@c] )
VAR common  = FILTER ( yearTbl, NOT ISBLANK ( [@p] ) )
VAR curC    = SUMX ( common, [@c] )
VAR prevC   = SUMX ( common, [@p] )
VAR commonTxt = CONCATENATEX ( common, FactData[Year], ", ", FactData[Year], ASC )
VAR newYrs    = CONCATENATEX ( FILTER ( yearTbl, ISBLANK ( [@p] ) ), FactData[Year], ", ", FactData[Year], ASC )

VAR compTbl =
    CALCULATETABLE (
        ADDCOLUMNS ( SUMMARIZE ( FactData, DimItem[Item Label] ),
            "@c", [Value Current],
            "@d", SUMX ( VALUES ( FactData[Year] ), [Delta CvP] ) ),
        REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ),
        FactData[Topic] = "Capital Needs", FactData[IsTotal] = FALSE () )
VAR nNeed    = COUNTROWS ( FILTER ( compTbl, [@c] > 0 ) )
VAR topRow   = TOPN ( 1, compTbl, [@c], DESC )
VAR topName  = MAXX ( topRow, DimItem[Item Label] )
VAR topShare = DIVIDE ( MAXX ( topRow, [@c] ), tot )
VAR chgRow   = TOPN ( 1, FILTER ( compTbl, [@d] <> 0 ), ABS ( [@d] ), DESC )
VAR chgName  = MAXX ( chgRow, DimItem[Item Label] )
VAR chgV     = MAXX ( chgRow, [@d] )

VAR pageText =
    "CAPITAL NEEDS " & UPPER ( period ) & nl & nl
    & IF ( tot = 0,
           "No capital needs planned in the current plan.",
           "Total capital needs (current plan): " & FORMAT ( tot, fmt ) & ", requested by " & nNeed
           & IF ( nNeed = 1, " company", " companies" )
           & IF ( y0 <> y1, "; peak in " & peakY & " (" & FORMAT ( peakV, fmt ) & ")", "" ) & "." & nl
           & "Largest need: " & topName & " (" & FORMAT ( topShare, "0%" ) & " of the total)." )
    & nl & nl
    & IF ( COUNTROWS ( common ) > 0,
           "Versus previous plan (" & commonTxt & "): " & FORMAT ( curC, fmt ) & " vs " & FORMAT ( prevC, fmt )
           & " (" & FORMAT ( curC - prevC, fmtD )
           & IF ( prevC <> 0, ", " & FORMAT ( DIVIDE ( curC - prevC, ABS ( prevC ) ), "+0%;-0%;0%" ), "" ) & ")"
           & IF ( curC > prevC, ": needs are higher than planned.",
                  IF ( curC < prevC, ": needs are lower than planned.", ": in line with the previous plan." ) )
           & IF ( NOT ISBLANK ( chgName ), nl & "Biggest change: " & chgName & " (" & FORMAT ( chgV, fmtD ) & ").", "" ),
           "No previous plan available for comparison." )
    & IF ( newYrs <> "" && COUNTROWS ( common ) > 0,
           nl & "Not in the previous plan: " & newYrs & " (new plan year).", "" )

-- one company selected
VAR lbl    = SELECTEDVALUE ( DimItem[Item Label] )
VAR rowYrs = ADDCOLUMNS ( yrs,
                 "@c", CALCULATE ( [Value Current], FactData[Topic] = "Capital Needs" ),
                 "@p", CALCULATE ( [Value Previous], FactData[Topic] = "Capital Needs" ) )
VAR rCur    = SUMX ( rowYrs, [@c] )
VAR rCommon = FILTER ( rowYrs, NOT ISBLANK ( [@p] ) )
VAR rD      = SUMX ( rCommon, [@c] - [@p] )
VAR rMaxRow = TOPN ( 1, rCommon, ABS ( [@c] - [@p] ), DESC )
VAR rMaxY   = MAXX ( rMaxRow, FactData[Year] )
VAR rMaxD   = MAXX ( rMaxRow, [@c] - [@p] )
VAR byYear  = CONCATENATEX ( rowYrs, FactData[Year] & ": " & FORMAT ( [@c] + 0, fmt ), " | ", FactData[Year], ASC )

VAR rowText =
    UPPER ( lbl ) & nl & nl
    & IF ( rCur = 0,
           "No capital needs planned for this company " & period & ".",
           IF ( y0 = y1,
                "Capital need in " & y1 & ": " & FORMAT ( rCur, fmt ) & " (" & FORMAT ( DIVIDE ( rCur, tot ), "0%" )
                & " of all needs in " & y1 & ").",
                "By year: " & byYear & "." & nl
                & "Total " & period & ": " & FORMAT ( rCur, fmt ) & " (" & FORMAT ( DIVIDE ( rCur, tot ), "0%" )
                & " of all capital needs)." ) )
    & IF ( COUNTROWS ( rCommon ) > 0,
           nl & "Versus previous plan: " & FORMAT ( rD, fmtD )
           & IF ( rD > 0, " (higher)", IF ( rD < 0, " (lower)", " (unchanged)" ) )
           & IF ( y0 <> y1 && rMaxD <> 0, "; largest change in " & rMaxY & " (" & FORMAT ( rMaxD, fmtD ) & ")", "" ) & ".",
           nl & "No previous plan for this company." )

RETURN IF ( HASONEVALUE ( DimItem[Item Label] ), rowText, pageText )


---------------------------------------------------------------------
 C) TOTAL REMITTANCE
 2 helper measures, then TR Comment.
 Uses the "Total ..." line if the sheet has one, otherwise the sum of lines.
---------------------------------------------------------------------

TR Cur =
VAR t = CALCULATE ( [Value Current], FactData[Topic] = "Remittance", FactData[IsTotal] = TRUE () )
RETURN IF ( ISBLANK ( t ), CALCULATE ( [Value Current], FactData[Topic] = "Remittance", FactData[IsTotal] = FALSE () ), t )

TR Prev =
VAR t = CALCULATE ( [Value Previous], FactData[Topic] = "Remittance", FactData[IsTotal] = TRUE () )
RETURN IF ( ISBLANK ( t ), CALCULATE ( [Value Previous], FactData[Topic] = "Remittance", FactData[IsTotal] = FALSE () ), t )


TR Comment =
VAR fmt  = "#,##0"
VAR fmtD = "+#,##0;-#,##0;0"
VAR nl   = UNICHAR ( 10 )

VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Remittance", NOT ISBLANK ( FactData[Value] ) )
VAR y0 = MINX ( yrs, FactData[Year] )
VAR y1 = MAXX ( yrs, FactData[Year] )
VAR period = IF ( y0 = y1, "in " & y1, y0 & "-" & y1 )

VAR yearTbl = ADDCOLUMNS ( yrs, "@c", ABS ( [TR Cur] ), "@p", ABS ( [TR Prev] ) )
VAR active  = FILTER ( yearTbl, [@c] <> 0 )
VAR nAct    = COUNTROWS ( active )
VAR tot     = SUMX ( yearTbl, [@c] )
VAR avgV    = AVERAGEX ( active, [@c] )
VAR peakRow = TOPN ( 1, active, [@c], DESC )
VAR peakY   = MAXX ( peakRow, FactData[Year] )
VAR peakV   = MAXX ( peakRow, [@c] )
VAR fY      = MINX ( active, FactData[Year] )
VAR lY      = MAXX ( active, FactData[Year] )
VAR fV      = MAXX ( FILTER ( active, FactData[Year] = fY ), [@c] )
VAR lV      = MAXX ( FILTER ( active, FactData[Year] = lY ), [@c] )
VAR common  = FILTER ( yearTbl, [@p] <> 0 )
VAR curC    = SUMX ( common, [@c] )
VAR prevC   = SUMX ( common, [@p] )
VAR commonTxt = CONCATENATEX ( common, FactData[Year], ", ", FactData[Year], ASC )
VAR dRow    = TOPN ( 1, common, ABS ( [@c] - [@p] ), DESC )
VAR dY      = MAXX ( dRow, FactData[Year] )
VAR dV      = MAXX ( dRow, [@c] - [@p] )
VAR zeroTxt = CONCATENATEX ( FILTER ( yearTbl, [@c] = 0 ), FactData[Year], ", ", FactData[Year], ASC )

-- single year (cell clicked)
VAR prevYear  = ABS ( CALCULATE ( [TR Cur], FactData[Year] = y1 - 1 ) )
VAR planTotal = SUMX ( ADDCOLUMNS ( CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( FactData[Year] ),
                    FactData[Topic] = "Remittance" ), "@a", ABS ( [TR Cur] ) ), [@a] )

VAR yearText =
    "TOTAL REMITTANCE " & y1 & nl & nl
    & IF ( tot = 0, "No remittance planned in " & y1 & ".",
           "Remittance to Group in " & y1 & ": " & FORMAT ( tot, fmt )
           & " (" & FORMAT ( DIVIDE ( tot, planTotal ), "0%" ) & " of the whole plan)."
           & IF ( prevYear <> 0, nl & "Versus " & ( y1 - 1 ) & ": " & FORMAT ( tot - prevYear, fmtD )
                  & " (" & FORMAT ( DIVIDE ( tot - prevYear, prevYear ), "+0%;-0%;0%" ) & ").", "" ) )
    & IF ( prevC <> 0,
           nl & "Versus previous plan: " & FORMAT ( curC - prevC, fmtD )
           & " (" & FORMAT ( DIVIDE ( curC - prevC, prevC ), "+0%;-0%;0%" ) & ").",
           nl & "No previous plan amount for " & y1 & "." )

VAR pageText =
    "TOTAL REMITTANCE " & UPPER ( period ) & nl & nl
    & IF ( tot = 0, "No remittance planned in the current plan.",
           "Remittance to Group (current plan): " & FORMAT ( tot, fmt ) & " over " & nAct
           & IF ( nAct = 1, " year", " years" ) & ", on average " & FORMAT ( avgV, fmt ) & " per year." & nl
           & "Highest in " & peakY & " (" & FORMAT ( peakV, fmt ) & ")"
           & IF ( fY <> lY,
                  "; " & lY & " vs " & fY & ": " & FORMAT ( DIVIDE ( lV - fV, fV ), "+0%;-0%;0%" ), "" ) & "." )
    & nl & nl
    & IF ( COUNTROWS ( common ) > 0,
           "Versus previous plan (" & commonTxt & "): " & FORMAT ( curC, fmt ) & " vs " & FORMAT ( prevC, fmt )
           & " (" & FORMAT ( curC - prevC, fmtD ) & ", " & FORMAT ( DIVIDE ( curC - prevC, prevC ), "+0%;-0%;0%" ) & ")."
           & IF ( dV <> 0, nl & "Largest change in " & dY & " (" & FORMAT ( dV, fmtD ) & ").", "" ),
           "The previous plan shows no amounts, so no comparison is possible. Please check the Previous block in the input file." )
    & IF ( zeroTxt <> "" && tot <> 0, nl & "No amount planned for: " & zeroTxt & ".", "" )

RETURN IF ( y0 = y1 && HASONEVALUE ( FactData[Year] ), yearText, pageText )
