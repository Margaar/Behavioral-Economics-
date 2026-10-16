HMS Invoices =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts",
    CONTAINSSTRING ( FactData[Item], "HMS invoices" ), NOT CONTAINSSTRING ( FactData[Item], "not eligible" ) )
HMS NotEligible =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts",
    CONTAINSSTRING ( FactData[Item], "not eligible" ) )
HMS Other =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "HMS Impacts",
    CONTAINSSTRING ( FactData[Item], "Other HMS" ) )
HMS Check = [HMS Invoices] + [HMS NotEligible] + [HMS Other] - [HMS Net]
HMS Share of Remittance =
DIVIDE ( [HMS Net],
    ABS ( CALCULATE ( [TR Cur], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Entity] ), REMOVEFILTERS ( DimEntity ) ) ) )
