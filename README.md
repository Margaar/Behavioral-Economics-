LK Comment All =
VAR gSel = ISFILTERED ( 'LK Guar Rows'[Row] ) && HASONEVALUE ( 'LK Guar Rows'[Row] )
VAR ent  = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR lbl  = SELECTEDVALUE ( 'LK Guar Rows'[Label] )
VAR kind = IF ( SELECTEDVALUE ( 'LK Guar Rows'[Kind] ) = "NB", "NB Guarantees – on NB Premiums",
                "IF Guarantees – on Mathematical Reserves" )
VAR isPct = SELECTEDVALUE ( 'LK Guar Rows'[Key] ) IN { "LINK", "AVG", "YR", "MR" }
VAR f   = IF ( isPct, "0.0%", "#,##0" )
VAR per = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Year], FactData[Period] ),
              REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
              FactData[Topic] = "Life KPIs", NOT ISBLANK ( FactData[Year] ) )
VAR yA  = MINX ( per, FactData[Year] )
VAR yZ  = MAXX ( per, FactData[Year] )
VAR series = CONCATENATEX ( per,
              FactData[Period] & " " & FORMAT ( CALCULATE ( [LK Guar Value], REMOVEFILTERS ( DimPeriod ) ), f ),
              ", ", FactData[Year], ASC )
VAR vA = CALCULATE ( [LK Guar Value], REMOVEFILTERS ( DimPeriod ),
             REMOVEFILTERS ( FactData[Period], FactData[Year] ), FactData[Year] = yA )
VAR vZ = CALCULATE ( [LK Guar Value], REMOVEFILTERS ( DimPeriod ),
             REMOVEFILTERS ( FactData[Period], FactData[Year] ), FactData[Year] = yZ )
VAR chg = IF ( isPct, FORMAT ( ( vZ - vA ) * 100, "+0.0;-0.0;0.0" ) & " pp",
                      FORMAT ( vZ - vA, "+#,##0;-#,##0;0" ) )
VAR gTxt = lbl & " (" & ent & ", " & kind & "): " & series & "."
           & IF ( NOT ISBLANK ( vA ) || NOT ISBLANK ( vZ ),
                  UNICHAR ( 10 ) & "Change " & yA & " to " & yZ & ": " & chg & "." )
RETURN IF ( gSel, gTxt, [LK Comment] )
