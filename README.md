SAA Comment =
VAR ent = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR one = HASONEVALUE ( FactData[SAA Part] )
VAR prt = IF ( one, SELECTEDVALUE ( FactData[SAA Part] ) )
VAR yA  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Actual" )
VAR yE  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Expected" )
VAR yS  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "SAA" )
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"

// ---------- page mode: whole company ----------
VAR act = [SAA MV Act]
VAR ex  = [SAA MV Exp]
VAR tg  = [SAA Tgt MV]
VAR pu  = [SAA Purch]
VAR sa  = [SAA Sales]
VAR re  = [SAA Redem]
VAR nb  = [SAA NetBal]
VAR inc = [SAA Income]
VAR dg  = [SAA Dur Gap]
VAR sP  = [SAA SCR Purch]
VAR sS  = [SAA SCR Sales]
VAR sN  = [SAA SCR Net]
VAR ror = [SAA RoRC]
VAR hasData = NOT ISBLANK ( act ) || NOT ISBLANK ( ex ) || NOT ISBLANK ( tg )
VAR cls = CALCULATETABLE (
              SUMMARIZE ( FactData, FactData[SAA Part], FactData[SAA Line] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Order] <= 9 )
VAR x   = ADDCOLUMNS ( cls, "@e", [SAA L Exp], "@t", [SAA L Tgt],
                            "@p", [SAA L Pct], "@lo", [SAA L Lo], "@hi", [SAA L Hi] )
VAR mv  = FILTER ( x, NOT ISBLANK ( [@t] ) || NOT ISBLANK ( [@e] ) )
VAR bigRow  = TOPN ( 1, mv, ABS ( [@t] - [@e] ), DESC )
VAR bigName = MAXX ( bigRow, FactData[SAA Line] & IF ( NOT one, " (" & FactData[SAA Part] & ")" ) )
VAR bigD    = MAXX ( bigRow, [@t] - [@e] )
VAR withPct = FILTER ( x, NOT ISBLANK ( [@p] ) )
VAR br  = FILTER ( withPct,
              ( NOT ISBLANK ( [@lo] ) && [@p] < [@lo] ) || ( NOT ISBLANK ( [@hi] ) && [@p] > [@hi] ) )
VAR nBr = COUNTROWS ( br )
VAR brTxt = CONCATENATEX ( br,
              FactData[SAA Line] & IF ( NOT one, " (" & FactData[SAA Part] & ")" ) & " "
              & FORMAT ( [@p], "0.0%" ) & " vs " & FORMAT ( [@lo], "0%" ) & "–" & FORMAT ( [@hi], "0%" ), "; " )
VAR pctSum = SUMX ( withPct, [@p] )
VAR flow = nb - ex
VAR gap  = nb - tg
VAR L1 = ent & IF ( one && prt <> "Total", " – " & prt ) & ": investments "
         & FORMAT ( act, fmt ) & " at Act " & yA & ", " & FORMAT ( ex, fmt ) & " expected for " & yE
         & ", SAA " & yS & " target " & FORMAT ( tg, fmt )
         & " (" & FORMAT ( tg - ex, fmtS ) & " vs " & yE & ")."
VAR L2 = IF ( NOT ISBLANK ( nb ),
         "Net balance " & yS & ": " & FORMAT ( flow, fmtS ) & " (purchases " & FORMAT ( pu, fmt )
         & ", sales " & FORMAT ( sa, fmt ) & ", redemptions/distributions " & FORMAT ( re, fmt )
         & "), MV after rebalancing " & FORMAT ( nb, fmt ) & "." )
VAR L3 = IF ( NOT ISBLANK ( bigD ) && bigD <> 0,
         "Largest reallocation: " & bigName & " " & FORMAT ( bigD, fmtS ) & " vs " & yE & "." )
VAR L4 = IF ( COUNTROWS ( withPct ) > 0,
         IF ( nBr = 0, "All target weights are within the SAA bounds.",
              nBr & IF ( nBr = 1, " asset class", " asset classes" ) & " outside the bounds: " & brTxt & "." ) )
VAR L5 = IF ( one && COUNTROWS ( withPct ) > 0 && ABS ( pctSum - 1 ) > 0.001,
         "Check: target weights add up to " & FORMAT ( pctSum, "0.0%" ) & ", not 100%." )
VAR L6 = IF ( NOT ISBLANK ( tg ) && NOT ISBLANK ( nb ) && ABS ( gap ) >= 1,
         "MV after net balance differs from the SAA target by " & FORMAT ( gap, fmtS )
         & " (market effect or input to check)." )
VAR L7 = IF ( NOT ISBLANK ( inc ),
         "Expected current income on purchases: " & FORMAT ( inc, "#,##0.0" ) & "." )
VAR L8 = IF ( one && NOT ISBLANK ( dg ), "Duration gap " & yE & ": " & FORMAT ( dg, "0.0" ) & " years." )
VAR L9 = IF ( NOT ISBLANK ( sP ) || NOT ISBLANK ( sS ),
         "SCR absorption " & yS & ": " & FORMAT ( sP, fmtS ) & " from purchases, "
         & FORMAT ( sS, fmtS ) & " from sales & redemptions (net " & FORMAT ( sP + sS, fmtS ) & ")"
         & IF ( NOT ISBLANK ( sN ), "; net SCR " & yE & ": " & FORMAT ( sN, fmt ) ) & "." )
VAR L10 = IF ( one && NOT ISBLANK ( ror ),
          "BU return on risk capital " & yS & ": " & FORMAT ( ror, "0.0%" ) & "." )
VAR lines = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ), ( 5, L5 ),
              ( 6, L6 ), ( 7, L7 ), ( 8, L8 ), ( 9, L9 ), ( 10, L10 ) }

// ---------- row mode: one asset class clicked in the matrix ----------
VAR rowSel = ISFILTERED ( FactData[SAA Line] ) && HASONEVALUE ( FactData[SAA Line] )
VAR ln   = SELECTEDVALUE ( FactData[SAA Line] )
VAR lAct = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", LEFT ( FactData[SAA Col], 6 ) = "MV Act" )
VAR lExp = [SAA L Exp]
VAR lTg  = [SAA L Tgt]
VAR lPct = [SAA L Pct]
VAR lLo  = [SAA L Lo]
VAR lHi  = [SAA L Hi]
VAR lNb  = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Net Balance" )
VAR lPu  = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases" )
VAR lSa  = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Sales" )
VAR lRe  = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Redemptions" )
VAR lYld = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Current Yield %" )
VAR lInc = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Current Income" )
VAR lPos = IF ( ISBLANK ( lPct ), "",
           IF ( NOT ISBLANK ( lLo ) && lPct < lLo, " – BELOW the lower bound",
           IF ( NOT ISBLANK ( lHi ) && lPct > lHi, " – ABOVE the upper bound", " – within the bounds" ) ) )
VAR R1 = ln & " (" & ent & IF ( one && prt <> "Total", " – " & prt ) & ")"
VAR R2 = IF ( CONTAINSSTRING ( ln, "Duration" ),
           FORMAT ( lAct, "0.0" ) & " years at Act " & yA & ", " & FORMAT ( lExp, "0.0" ) & " expected for " & yE
           & IF ( NOT ISBLANK ( lPct ), "; SAA " & yS & " target " & FORMAT ( lPct, "0.0" ) ) & ".",
           "MV " & FORMAT ( lAct, fmt ) & " at Act " & yA & ", " & FORMAT ( lExp, fmt ) & " expected for " & yE
           & "; after net balance " & yS & ": " & FORMAT ( lNb, fmt ) & " (" & FORMAT ( lNb - lExp, fmtS ) & ")." )
VAR R3 = IF ( NOT ISBLANK ( lPu ) || NOT ISBLANK ( lSa ) || NOT ISBLANK ( lRe ),
           "Flows " & yS & ": purchases " & FORMAT ( lPu, fmt ) & ", sales " & FORMAT ( lSa, fmt )
           & ", redemptions/distributions " & FORMAT ( lRe, fmt ) & "." )
VAR R4 = IF ( NOT ISBLANK ( lTg ) && NOT CONTAINSSTRING ( ln, "Duration" ),
           "SAA " & yS & " target " & FORMAT ( lTg, fmt ) & " = " & FORMAT ( lPct, "0.0%" )
           & IF ( NOT ISBLANK ( lLo ) || NOT ISBLANK ( lHi ),
                  " (bounds " & FORMAT ( lLo, "0%" ) & "–" & FORMAT ( lHi, "0%" ) & ")" ) & lPos & "." )
VAR R5 = IF ( NOT ISBLANK ( lYld ) && lYld <> 0,
           "Purchases at " & FORMAT ( lYld, "0.0%" ) & " current yield"
           & IF ( NOT ISBLANK ( lInc ), ", expected current income " & FORMAT ( lInc, "#,##0.0" ) ) & "." )
VAR rowLines = { ( 1, R1 ), ( 2, R2 ), ( 3, R3 ), ( 4, R4 ), ( 5, R5 ) }

RETURN
IF ( rowSel,
     CONCATENATEX ( FILTER ( rowLines, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ),
     IF ( NOT hasData, ent & ": no SAA data filled in.",
          CONCATENATEX ( FILTER ( lines, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) ) )
