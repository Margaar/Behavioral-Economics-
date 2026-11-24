    Countries = [AT="Austria", BE="Belgium", BG="Bulgaria", CH="Switzerland", CS="Serbia", RS="Serbia",
                 CZ="Czech Republic", DE="Germany", ES="Spain", FR="France", GR="Greece", HR="Croatia",
                 HU="Hungary", IE="Ireland", IT="Italy", LI="Liechtenstein", LU="Luxembourg", MC="Monaco",
                 ME="Montenegro", NL="Netherlands", PL="Poland", PT="Portugal", RO="Romania", SI="Slovenia",
                 SK="Slovakia", GB="United Kingdom", AR="Argentina", BR="Brazil", CL="Chile", EC="Ecuador",
                 MX="Mexico", CO="Colombia", MY="Malaysia", PH="Philippines", IN="India", CN="China",
                 HK="Hong Kong", TH="Thailand", ID="Indonesia", US="United States", AO="Angola", TR="Turkey"],
    Ccy       = [CH="CHF", LI="CHF", CS="RSD", RS="RSD", CZ="CZK", HU="HUF", PL="PLN", RO="RON",
                 GB="GBP", AR="ARS", BR="BRL", CL="CLP", EC="USD", MX="MXN", CO="COP", MY="MYR",
                 PH="PHP", IN="INR", CN="CNY", HK="HKD", TH="THB", ID="IDR", US="USD", AO="AOA", TR="TRY"],
    SII       = {"AT","BE","BG","HR","CY","CZ","DK","EE","FI","FR","DE","GR","HU","IE","IT","LV","LT",
                 "LU","MT","NL","PL","PT","RO","SK","SI","ES","SE","NO","IS","LI"},

    WithCtry  = Table.AddColumn(WithCC, "Country", each Record.FieldOrDefault(Countries, [Country Code], [Country Code]), type text),
    WithCcy   = Table.AddColumn(WithCtry, "Currency", each Record.FieldOrDefault(Ccy, [Country Code], "EUR"), type text),
    Final     = Table.AddColumn(WithCcy, "Regime", each
                  if List.Contains(SII, [Country Code]) then "Solvency II"
                  else if List.Contains({"CH"}, [Country Code]) then "SST (CH)"
                  else "Non-SII", type text)
