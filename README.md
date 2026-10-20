LK Comment =
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR ent  = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR rowSel = ISFILTERED ( FactData[LK Line] ) && HASONEVALUE ( FactData[LK Line] )
VAR yA = CALCULATE ( MIN ( FactData[Year] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs" )
VAR yZ = CALCULATE ( MAX ( FactData[Year] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs" )
VAR pA = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yA )
VAR pZ = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yZ )
-- page mode
VAR vA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "VNP" ), FactData[Year] = yA )
VAR vZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "VNP" ), FactData[Year] = yZ )
VAR nA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "Total NBV" ), FactData[Year] = yA )
VAR nZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "Total NBV" ), FactData[Year] = yZ )
VAR lob  = CALCULATETABLE ( VALUES ( FactData[LK Line] ),
               REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
               FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[LK Order] <= 3 )
VAR lobX = FILTER ( ADDCOLUMNS ( lob, "@v",
               CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
                   FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[Year] = yZ ) ),
               NOT ISBLANK ( [@v] ) && [@v] <> 0 )
VAR lobT = TOPN ( 1, lobX, [@v], DESC )
VAR lobN = MAXX ( lobT, FactData[LK Line] )
VAR lobV = MAXX ( lobT, [@v] )
VAR rec  = CALCULATETABLE ( VALUES ( FactData[LK Line] ),
               REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
               FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[LK Order] > 3,
               NOT CONTAINSSTRING ( FactData[Item], "Total" ) )
VAR recX = FILTER ( ADDCOLUMNS ( rec, "@v",
               CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
                   FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[Year] = yZ ) ),
               NOT ISBLANK ( [@v] ) && [@v] <> 0 )
VAR recT = TOPN ( 1, recX, ABS ( [@v] ), DESC )
VAR recN = MAXX ( recT, FactData[LK Line] )
VAR recV = MAXX ( recT, [@v] )
VAR pvZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "PVNBP",
             CONTAINSSTRING ( FactData[Item], "Total" ), FactData[Year] = yZ )
VAR mgZ = DIVIDE ( nZ, pvZ )
VAR gNB = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Guarantee-linked NB" ), FactData[Year] = yZ )
VAR aNB = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Average Guarantee NB" ), FactData[Year] = yZ )
VAR gIF = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Guarantee-linked Mathematical" ), FactData[Year] = yZ )
VAR aIF = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[LK Line], FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Average Guarantee on Mathematical" ), FactData[Year] = yZ )
VAR L1 = ent & ": Total Life SII VNP (net of tax) " & FORMAT ( vA, fmt ) & " in " & pA & ", "
         & FORMAT ( vZ, fmt ) & " in " & pZ & " (" & FORMAT ( vZ - vA, fmtS ) & ")."
VAR L2 = IF ( NOT ISBLANK ( nA ) || NOT ISBLANK ( nZ ),
         "NBV official (gross of tax): " & FORMAT ( nA, fmt ) & " in " & pA & ", " & FORMAT ( nZ, fmt ) & " in " & pZ & "." )
VAR L3 = IF ( NOT ISBLANK ( lobV ), "Largest line of business in " & pZ & ": " & lobN & " (" & FORMAT ( lobV, fmt ) & ")." )
VAR L4 = IF ( NOT ISBLANK ( recV ), "Main item from NBV to VNP in " & pZ & ": " & recN & " " & FORMAT ( recV, fmtS ) & "." )
VAR L5 = IF ( NOT ISBLANK ( mgZ ) && mgZ <> 0, "NBV / PVNBP margin " & pZ & ": " & FORMAT ( mgZ, "0.0%" ) & "." )
VAR L6 = IF ( NOT ISBLANK ( gNB ) && gNB <> 0, "Guarantee-linked NB premiums " & pZ & ": " & FORMAT ( gNB, "0%" )
         & IF ( NOT ISBLANK ( aNB ) && aNB <> 0, ", average guarantee " & FORMAT ( aNB, "0.00%" ) ) & "." )
VAR L7 = IF ( NOT ISBLANK ( gIF ) && gIF <> 0, "Guarantee-linked mathematical reserves " & pZ & ": " & FORMAT ( gIF, "0%" )
         & IF ( NOT ISBLANK ( aIF ) && aIF <> 0, ", average guarantee " & FORMAT ( aIF, "0.00%" ) ) & "." )
-- row mode: one line clicked
VAR rl    = SELECTEDVALUE ( FactData[LK Line] )
VAR rTbl  = CALCULATE ( MAX ( FactData[LK Table] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
                FactData[Topic] = "Life KPIs" )
VAR isPct = CONTAINSSTRING ( rl, "(%)" ) || CONTAINSSTRING ( rl, "Guarantee-linked" )
            || CONTAINSSTRING ( rl, "Average Guarantee" )
VAR rf    = IF ( isPct, "0.0%", fmt )
VAR per   = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Year], FactData[Period] ),
                REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ), FactData[Topic] = "Life KPIs" )
VAR series = CONCATENATEX ( per, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ) ), rf ),
                ", ", FactData[Year], ASC )
VAR rA = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yA )
VAR rZ = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yZ )
VAR R1 = rl & " (" & ent & ", " & rTbl & "): " & series & "."
VAR R2 = IF ( NOT ISBLANK ( rA ) || NOT ISBLANK ( rZ ),
         "Change " & pA & " to " & pZ & ": "
         & IF ( isPct, FORMAT ( ( rZ - rA ) * 100, "+0.0;-0.0;0.0" ) & " pp", FORMAT ( rZ - rA, fmtS ) ) & "." )
VAR pg = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ), ( 5, L5 ), ( 6, L6 ), ( 7, L7 ) }
RETURN
IF ( rowSel, R1 & UNICHAR ( 10 ) & R2,
     CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )
