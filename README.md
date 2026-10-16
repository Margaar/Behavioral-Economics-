HMS Share of Remittance =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts" )
VAR tr  = ABS ( CALCULATE ( [TR Cur], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Entity] ), REMOVEFILTERS ( DimEntity ),
              TREATAS ( yrs, FactData[Year] ) ) )
RETURN DIVIDE ( [HMS Net], tr )



VAR rowVal = CALCULATE ( [Value Current], FactData[Topic] = "HMS Impacts" )
VAR rowText =
    UPPER ( lbl ) & ": " & who & nl & nl
    & IF ( y0 = y1,
           "Value in " & y1 & ": " & FORMAT ( rowVal + 0, fmt )
           & IF ( net <> 0, " (" & FORMAT ( DIVIDE ( rowVal, net ), "0%" ) & " of the total net cash impact in " & y1 & ")", "" ) & ".",
           "By year: " & byYear & "." & nl
           & "Total " & period & ": " & FORMAT ( rowVal + 0, fmt )
           & IF ( net <> 0, " (" & FORMAT ( DIVIDE ( rowVal, net ), "0%" ) & " of the total net cash impact)", "" ) & "." )
