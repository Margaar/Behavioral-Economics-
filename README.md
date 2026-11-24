    WithReg   = Table.AddColumn(WithCcy, "Regime", each
                  if [Legal Entity Code] = "FR158" then "Local Solvency (non-SII)"
                  else if List.Contains(SII, [Country Code]) then "Solvency II"
                  else if [Country Code] = "CH" then "SST (CH)"
                  else "Non-SII", type text),
