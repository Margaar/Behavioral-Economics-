SAA PH Part =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
    FactData[Topic] = "SAA", CONTAINSSTRING ( FactData[Item], "Participation" ) )



SAA Tax Rate =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
    FactData[Topic] = "SAA", CONTAINSSTRING ( FactData[Item], "Tax Rate" ) )


SAA c Net Income on SCR =
VAR inc = CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
              FactData[SAA Col] = "Current Income", FactData[SAA Order] = 10 )
VAR scr = CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "SCR",
              LEFT ( FactData[SAA Col], 8 ) = "SCR Purc", FactData[SAA Order] = 10 )
VAR ph  = [SAA PH Part]
VAR tax = [SAA Tax Rate]
RETURN DIVIDE ( inc * ( 1 - ph ) * ( 1 - tax ), scr )



SAA c Net TR on SCR =
VAR tr  = CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
              FactData[SAA Col] = "Total Return", FactData[SAA Order] = 10 )
VAR mv  = CALCULATE ( SUM ( FactData[Value] ),        // <- "total mv (eur m)": см. вопрос ниже
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
              FactData[SAA Col] = "Purchases", FactData[SAA Order] = 10 )
VAR scr = CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "SCR",
              LEFT ( FactData[SAA Col], 8 ) = "SCR Purc", FactData[SAA Order] = 10 )
VAR ph  = [SAA PH Part]
VAR tax = [SAA Tax Rate]
RETURN DIVIDE ( tr * ( 1 - ph ) * ( 1 - tax ) + mv * ( 1 - tax ), scr )


