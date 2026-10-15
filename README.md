VAR st =
    IF ( ISBLANK ( r1 ), "cannot be calculated (SCR missing)",
    IF ( r1 < hardLim, "BELOW the " & FORMAT ( hardLim, "0%" ) & " threshold: action needed",
    IF ( r1 < softLim, "between the " & FORMAT ( hardLim, "0%" ) & " and " & FORMAT ( softLim, "0%" ) & " thresholds: limited buffer",
         "above the " & FORMAT ( softLim, "0%" ) & " threshold: comfortable buffer" ) ) )
