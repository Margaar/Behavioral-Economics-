HMS Comment =
VAR fmt  = "#,##0"
VAR fmtD = "+#,##0;-#,##0;0"
VAR nl   = UNICHAR ( 10 )
VAR yrs  = CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts" )
VAR y0   = MINX ( yrs, FactData[Year] )
VAR y1   = MAXX ( yrs, FactData[Year] )
VAR period = IF ( y0 = y1, "in " & y1, y0 & "-" & y1 )
VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts" )
VAR who  = IF ( nEnt = 1,
               CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts" ),
               nEnt & " companies" )
VAR net = [HMS Net]
VAR inv = [HMS Invoices]
VAR ne  = [HMS NotEligible]
VAR oth = [HMS Other]
VAR chk = ROUND ( [HMS Check], 0 )
VAR tr  = ABS ( CALCULATE ( [TR Cur], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Entity] ), REMOVEFILTERS ( DimEntity ),
              FactData[Year] >= y0, FactData[Year] <= y1 ) )
VAR share = DIVIDE ( net, tr )
VAR yearTbl = ADDCOLUMNS ( yrs, "@n", [HMS Net] )
VAR peakRow = TOPN ( 1, FILTER ( yearTbl, [@n] <> 0 ), ABS ( [@n] ), DESC )
VAR compTbl =
    CALCULATETABLE ( ADDCOLUMNS ( SUMMARIZE ( FactData, FactData[Entity] ), "@n", [HMS Net] ),
        REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts" )
VAR topRow = TOPN ( 1, FILTER ( compTbl, [@n] <> 0 ), ABS ( [@n] ), DESC )
VAR pageText =
    "HMS IMPACT ON REMITTANCE (CASH VIEW): " & UPPER ( who ) & ", " & period & nl & nl
    & IF ( net = 0 && inv = 0 && ne = 0 && oth = 0,
           "No HMS amounts have been entered for this selection.",
           "Total net cash impact to be neutralised: " & FORMAT ( net, fmt )
           & IF ( y0 <> y1, "; highest in " & MAXX ( peakRow, FactData[Year] ) & " (" & FORMAT ( MAXX ( peakRow, [@n] ), fmt ) & ")", "" ) & "." & nl
           & "Made of: AG HMS invoices " & FORMAT ( inv + 0, fmt )
           & ", not eligible for neutralisation " & FORMAT ( ne + 0, fmt )
           & ", other HMS cash impacts " & FORMAT ( oth + 0, fmt ) & "." & nl
           & IF ( tr <> 0, "This equals " & FORMAT ( share, "0.0%" ) & " of the country remittance " & period & "." & nl, "" )
           & IF ( nEnt > 1 && NOT ISEMPTY ( topRow ),
                  "Largest contributor: " & MAXX ( topRow, FactData[Entity] ) & " ("
                  & FORMAT ( DIVIDE ( MAXX ( topRow, [@n] ), net ), "0%" ) & " of the total)." & nl, "" )
           & nl
           & IF ( chk = 0, "Check: components reconcile with the Total Net Cash impact.",
                  "Check: components differ from the total by " & FORMAT ( chk, fmtD ) & ". Please review the input file." ) )
    & nl & "Only amounts already agreed with Group Capital Steering are included."
VAR lbl = SELECTEDVALUE ( DimItem[Item Label] )
VAR byYear = CONCATENATEX ( yrs, FactData[Year] & ": " & FORMAT ( CALCULATE ( [Value Current], FactData[Topic] = "HMS Impacts" ) + 0, fmt ),
                 " | ", FactData[Year], ASC )
VAR rowText =
    UPPER ( lbl ) & ": " & who & nl & nl
    & "By year: " & byYear & "." & nl
    & "Total " & period & ": " & FORMAT ( CALCULATE ( [Value Current], FactData[Topic] = "HMS Impacts" ) + 0, fmt ) & "."
RETURN IF ( HASONEVALUE ( DimItem[Item Label] ), rowText, pageText )
