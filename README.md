SOL EOF =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Eligible Own Funds" ) )
SOL SCR =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Solvency Capital Requirement" ) )
SOL FTC =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Free Tangible Capital" ) )
SOL NetResult =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Local GAAP Net Result" ) )
SOL Distributions =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Total Distributions" ) )
SOL OtherDistr =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Other distributions" ) )
SOL ForeseeableDiv =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Foreseeable" ) )
SOL CapNeeds =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Capital Needs" ) )
