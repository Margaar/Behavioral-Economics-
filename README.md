    Final    = Table.RenameColumns(Table.RemoveColumns(Scen, {"Scenario", "FIdx", "FirstIdx", "Cnt"}),
                   {{"Scen", "Scenario"}}),
    NoEntCountry = Table.ReplaceValue(Final, each [Entity], each
                   if List.Contains({"Remittance", "Remittance Holding", "Capital Needs", "HMS Impacts"}, [Topic])
                   then null else [Entity], Replacer.ReplaceValue, {"Entity"})
in
    NoEntCountry

        





VAR it = SELECTEDVALUE ( FactData[LK Line] )
RETURN IF ( CONTAINSSTRING ( it, "(%)" ) || CONTAINSSTRING ( it, "Guarantee-linked" )
            || CONTAINSSTRING ( it, "Average Guarantee" ), "0.0%", "#,##0" )
        
