RH BoP =
VAR y0 = CALCULATE ( MIN ( FactData[Year] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding" )
RETURN
    CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding",
        CONTAINSSTRING ( FactData[Item], "(BoP)" ), FactData[Year] = y0 )

______________


RH EoP =
VAR y1 = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding" )
RETURN
    CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding",
        CONTAINSSTRING ( FactData[Item], "(EoP)" ), FactData[Year] = y1 )

_____________


RH Div In =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding",
    CONTAINSSTRING ( FactData[Item], "Dividends inflow" ) )


_________

RH Div Out =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding",
    CONTAINSSTRING ( FactData[Item], "Dividends outflow" ) )

___________

RH Other Flows =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding",
    NOT ( LEFT ( FactData[Item], 3 ) = "o/w"
          || CONTAINSSTRING ( FactData[Item], "(BoP)" )
          || CONTAINSSTRING ( FactData[Item], "(EoP)" )
          || CONTAINSSTRING ( FactData[Item], "Net Result" )
          || CONTAINSSTRING ( FactData[Item], "Dividends" ) ) )

___________


RH Check = [RH BoP] + [RH Div In] + [RH Div Out] + [RH Other Flows] - [RH EoP]



____

RH Comment =
VAR fmt  = "#,##0"
VAR nl   = UNICHAR ( 10 )
VAR y0   = CALCULATE ( MIN ( FactData[Year] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding" )
VAR y1   = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding" )
VAR period = IF ( y0 = y1, "in " & y0, "over " & y0 & "-" & y1 )
VAR bop  = [RH BoP]
VAR eop  = [RH EoP]
VAR dIn  = [RH Div In]
VAR dOut = [RH Div Out]
VAR oth  = [RH Other Flows]
VAR chk  = [RH Check]
VAR cover = DIVIDE ( dIn, -dOut )

-- largest contributing subsidiary
VAR subs =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( DimItem[Item Label] ), "@v", [Value Current] ),
        REMOVEFILTERS ( DimItem ), DimItem[IsSubItem] = TRUE (), FactData[Topic] = "Remittance Holding" )
VAR topRow   = TOPN ( 1, subs, [@v], DESC )
VAR topName  = SUBSTITUTE ( MAXX ( topRow, DimItem[Item Label] ), "o/w from: ", "" )
VAR topShare = DIVIDE ( MAXX ( topRow, [@v] ), dIn )

-- lowest year-end cash in the period
VAR eopYears =
    ADDCOLUMNS (
        CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding" ),
        "@e", CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Remittance Holding",
                          CONTAINSSTRING ( FactData[Item], "(EoP)" ) ) )
VAR lowRow  = TOPN ( 1, eopYears, [@e], ASC )
VAR lowYear = MAXX ( lowRow, FactData[Year] )
VAR lowVal  = MAXX ( lowRow, [@e] )

VAR pageText =
    "REMITTANCE AT HOLDING LEVEL " & UPPER ( period ) & nl & nl
    & "Cash position: " & FORMAT ( bop, fmt ) & " at the start, " & FORMAT ( eop, fmt ) & " at the end ("
    & IF ( eop >= bop, "+", "" ) & FORMAT ( eop - bop, fmt ) & ")." & nl
    & "Dividends received from subsidiaries: " & FORMAT ( dIn, fmt )
    & IF ( dIn <> 0 && NOT ISBLANK ( topName ),
           "; largest contributor " & topName & " (" & FORMAT ( topShare, "0%" ) & ")", "" ) & "." & nl
    & "Dividends paid to shareholders: " & FORMAT ( -dOut, fmt ) & ". "
    & IF ( ISBLANK ( cover ), "",
           "Subsidiary dividends cover " & FORMAT ( cover, "0%" ) & " of distributions"
           & IF ( cover >= 1, ": distributions are fully funded by remittances.",
                  ": " & FORMAT ( -dOut - dIn, fmt ) & " is funded from existing cash." ) ) & nl
    & "Other cash flows (interest, capital, expenses, other): " & FORMAT ( oth, fmt ) & "."
    & IF ( y0 <> y1, nl & "Lowest year-end cash: " & FORMAT ( lowVal, fmt ) & " in " & lowYear & ".", "" )
    & nl & nl
    & IF ( ROUND ( chk, 0 ) = 0,
           "Check: the cash bridge reconciles with the reported Cash Position (EoP).",
           "Check: calculated and reported EoP differ by " & FORMAT ( chk, fmt ) & ". Please review the input file." )

-- comment for one selected row
VAR isRow   = ISFILTERED ( DimItem[Item Label] ) && HASONEVALUE ( DimItem[Item Label] )
VAR lbl     = SELECTEDVALUE ( DimItem[Item Label] )
VAR it      = SELECTEDVALUE ( DimItem[Item] )
VAR v       = [Value Current]
VAR isSub   = LEFT ( it, 3 ) = "o/w"
VAR isStock = CONTAINSSTRING ( it, "(BoP)" ) || CONTAINSSTRING ( it, "(EoP)" )
VAR prevV   = CALCULATE ( [Value Current], FactData[Year] = y0 - 1 )
VAR byYear  = CONCATENATEX ( VALUES ( FactData[Year] ),
                  FactData[Year] & ": " & FORMAT ( [Value Current], fmt ), " | ", FactData[Year], ASC )

VAR rowText =
    UPPER ( lbl ) & nl & nl
    & IF ( y0 = y1,
           "Value in " & y0 & ": " & FORMAT ( v, fmt )
           & IF ( NOT ISBLANK ( prevV ),
                  " (" & IF ( v >= prevV, "+", "" ) & FORMAT ( v - prevV, fmt ) & " vs " & ( y0 - 1 ) & ")", "" ) & ".",
           "By year: " & byYear & "."
           & IF ( NOT isStock, nl & "Total " & period & ": " & FORMAT ( v, fmt ) & ".", "" ) )
    & IF ( isSub && dIn <> 0,
           nl & "Share of dividends received " & period & ": " & FORMAT ( DIVIDE ( v, dIn ), "0%" ) & ".", "" )
    & IF ( ISBLANK ( v ) || v = 0, nl & "No amounts planned for this line.", "" )

RETURN IF ( isRow, rowText, pageText )

___________

