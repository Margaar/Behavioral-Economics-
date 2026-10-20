LK Guar Rows =
DATATABLE (
    "Kind", STRING, "Order", INTEGER, "Label", STRING, "Key", STRING,
    {
        { "NB", 1, "Total NB Premiums",                           "BASE" },
        { "NB", 2, "Total guaranteed NB Premiums",                "TOTG" },
        { "NB", 3, "On yearly basis guaranteed",                  "YA"   },
        { "NB", 4, "At maturity guaranteed",                      "MA"   },
        { "NB", 5, "Guarantee-linked NB Premiums (%)",            "LINK" },
        { "NB", 6, "Average Guarantee NB Premiums (%)",           "AVG"  },
        { "NB", 7, "On yearly basis guarantee (%)",               "YR"   },
        { "NB", 8, "At maturity guarantee (%)",                   "MR"   },
        { "IF", 1, "Total Mathematical Reserves",                 "BASE" },
        { "IF", 2, "Total guaranteed Mathematical Reserves",      "TOTG" },
        { "IF", 3, "On yearly basis guaranteed",                  "YA"   },
        { "IF", 4, "At maturity guaranteed",                      "MA"   },
        { "IF", 5, "Guarantee-linked Mathematical Reserves (%)",  "LINK" },
        { "IF", 6, "Average Guarantee on Mathematical Reserves (%)", "AVG" },
        { "IF", 7, "On yearly basis guarantee (%)",               "YR"   },
        { "IF", 8, "At maturity guarantee (%)",                   "MR"   }
    } )






    LK Guar Value =
VAR k    = SELECTEDVALUE ( 'LK Guar Rows'[Key] )
VAR tbl  = IF ( SELECTEDVALUE ( 'LK Guar Rows'[Kind] ) = "NB", "NB Guarantees", "IF Guarantees" )
VAR base = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = tbl,
               LEFT ( FactData[Item], 5 ) = "Total" && NOT CONTAINSSTRING ( FactData[Item], "guaranteed" ) )
VAR totIn = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = tbl,
               LEFT ( FactData[Item], 5 ) = "Total" && CONTAINSSTRING ( FactData[Item], "guaranteed" ) )
VAR yA = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = tbl,
               CONTAINSSTRING ( FactData[Item], "On yearly" ) && CONTAINSSTRING ( FactData[Item], "guaranteed" ) )
VAR mA = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = tbl,
               CONTAINSSTRING ( FactData[Item], "At maturity" ) && CONTAINSSTRING ( FactData[Item], "guaranteed" ) )
VAR yR = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = tbl,
               CONTAINSSTRING ( FactData[Item], "On yearly" ) && CONTAINSSTRING ( FactData[Item], "(%)" ) )
VAR mR = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = tbl,
               CONTAINSSTRING ( FactData[Item], "At maturity" ) && CONTAINSSTRING ( FactData[Item], "(%)" ) )
VAR avgIn = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = tbl,
               CONTAINSSTRING ( FactData[Item], "Average Guarantee" ) )
VAR totG = IF ( NOT ISBLANK ( totIn ) && totIn <> 0, totIn, COALESCE ( yA, 0 ) + COALESCE ( mA, 0 ) )
VAR avg  = IF ( NOT ISBLANK ( avgIn ) && avgIn <> 0, avgIn,
               DIVIDE ( COALESCE ( yA, 0 ) * COALESCE ( yR, 0 ) + COALESCE ( mA, 0 ) * COALESCE ( mR, 0 ),
                        COALESCE ( yA, 0 ) + COALESCE ( mA, 0 ) ) )
RETURN
SWITCH ( k,
    "BASE", base,
    "TOTG", IF ( totG <> 0, totG ),
    "YA",   yA,
    "MA",   mA,
    "LINK", DIVIDE ( totG, base ),
    "AVG",  IF ( avg <> 0, avg ),
    "YR",   yR,
    "MR",   mR )
