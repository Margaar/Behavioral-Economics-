SAA c Yield % =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
VAR t = ADDCOLUMNS (
            CALCULATETABLE ( VALUES ( FactData[SAA Part] ), REMOVEFILTERS ( FactData[SAA Part] ),
                FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation" ),
            "@p", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases" ),
            "@y", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Current Yield %" ) )
VAR pSum = SUMX ( t, [@p] )
RETURN
IF ( LEFT ( lineNm, 5 ) = "Total", BLANK (),
     IF ( pSum <> 0, SUMX ( t, [@p] * [@y] ) / pSum, MAXX ( t, [@y] ) ) )




     SAA c Total Return % =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
VAR t = ADDCOLUMNS (
            CALCULATETABLE ( VALUES ( FactData[SAA Part] ), REMOVEFILTERS ( FactData[SAA Part] ),
                FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation" ),
            "@p", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases" ),
            "@y", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Total Return Yield %" ) )
VAR pSum = SUMX ( t, [@p] )
RETURN
IF ( LEFT ( lineNm, 5 ) = "Total", BLANK (),
     IF ( pSum <> 0, SUMX ( t, [@p] * [@y] ) / pSum, MAXX ( t, [@y] ) ) )
