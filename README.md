NL Comment =
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR ent  = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR rowSel = ISFILTERED ( FactData[NL Line] ) && HASONEVALUE ( FactData[NL Line] )
VAR yA = CALCULATE ( MIN ( FactData[Year] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC" )
VAR yZ = CALCULATE ( MAX ( FactData[Year] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC" )
VAR pA = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC", FactData[Year] = yA )
VAR pZ = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC", FactData[Year] = yZ )
VAR epA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yA )
VAR epZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yZ )
VAR epG = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Gross", FactData[Year] = yZ )
VAR moZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             FactData[Item] = "Motor", FactData[Measure] = "Net", FactData[Year] = yZ )
VAR crA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "CoR",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yA )
VAR crZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "CoR",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yZ )
VAR siA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "SII CY Technical Result", FactData[Year] = yA )
VAR siZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "SII CY Technical Result", FactData[Year] = yZ )
VAR roA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "RoRC", CONTAINSSTRING ( FactData[Item], "RoRC" ),
             FactData[Measure] = "Net", FactData[PeriodType] = "Actual" )
VAR roB = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[NL Line], FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "RoRC", CONTAINSSTRING ( FactData[Item], "RoRC" ),
             FactData[Measure] = "Net", FactData[PeriodType] = "Benchmark" )
VAR cagr = IF ( epA > 0 && epZ > 0 && yZ > yA, ( epZ / epA ) ^ ( 1 / ( yZ - yA ) ) - 1 )
VAR N1 = IF ( NOT ISBLANK ( epZ ) && epZ <> 0,
         ent & ": net earned premiums (excl. GC&C) " & FORMAT ( epA, fmt ) & " in " & pA & ", "
         & FORMAT ( epZ, fmt ) & " in " & pZ & IF ( NOT ISBLANK ( cagr ), " (CAGR " & FORMAT ( cagr, "+0.0%;-0.0%" ) & ")" ) & "." )
VAR N2 = IF ( NOT ISBLANK ( epZ ) && epZ <> 0 && NOT ISBLANK ( moZ ),
         "Motor " & FORMAT ( DIVIDE ( moZ, epZ ), "0%" ) & " of net premiums in " & pZ
         & IF ( NOT ISBLANK ( epG ) && epG <> 0, "; retention (net / gross) " & FORMAT ( DIVIDE ( epZ, epG ), "0%" ) ) & "." )
VAR N3 = IF ( NOT ISBLANK ( crZ ) && crZ <> 0,
         "Net CoR (excl. GC&C): " & FORMAT ( crA, "0.0%" ) & " in " & pA & ", " & FORMAT ( crZ, "0.0%" ) & " in " & pZ
         & IF ( crZ < 1, " – below 100% (underwriting profit).", " – above 100% (underwriting loss)." ) )
VAR N4 = IF ( NOT ISBLANK ( siZ ) || NOT ISBLANK ( siA ),
         "Non-life SII CY technical result: " & FORMAT ( siA, fmt ) & " in " & pA & ", " & FORMAT ( siZ, fmt ) & " in " & pZ
         & " (" & FORMAT ( siZ - siA, fmtS ) & ")." )
VAR N5 = IF ( NOT ISBLANK ( roA ) && roA <> 0,
         "RoRC (excl. GC&C, net) " & FORMAT ( roA, "0.0%" )
         & IF ( NOT ISBLANK ( roB ) && roB <> 0, " vs benchmark " & FORMAT ( roB, "0.0%" )
                & IF ( roA >= roB, " – above benchmark.", " – below benchmark." ), "." ) )
VAR rl  = SELECTEDVALUE ( FactData[NL Line] )
VAR rG  = CALCULATE ( MAX ( FactData[NL Group] ), REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year], FactData[Measure] ),
              FactData[Topic] = "Non-Life KPIs" )
VAR rf  = IF ( rG IN { "CoR", "RoRC" }, "0.0%", fmt )
VAR per = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[PeriodType], FactData[Year], FactData[Period] ),
              REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year], FactData[Measure] ), FactData[Topic] = "Non-Life KPIs" )
VAR sN  = CONCATENATEX ( per, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Measure] ), FactData[Measure] = "Net" ), rf ), ", ", FactData[Year], ASC )
VAR sG  = CONCATENATEX ( per, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Measure] ), FactData[Measure] = "Gross" ), rf ), ", ", FactData[Year], ASC )
VAR R1 = rl & " (" & ent & ")" & UNICHAR ( 10 ) & "Net: " & sN & "." & UNICHAR ( 10 ) & "Gross: " & sG & "."
VAR pg = { ( 1, N1 ), ( 2, N2 ), ( 3, N3 ), ( 4, N4 ), ( 5, N5 ) }
VAR txt = CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC )
RETURN
IF ( rowSel, R1, IF ( txt = "", ent & ": no Non-Life KPI data filled in.", txt ) )
