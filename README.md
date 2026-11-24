    Dist      = Table.Distinct(NoBlank, {"Legal Entity Code"}),

    // убрать кавычки в начале названий
    Names     = Table.TransformColumns(Dist, {{"Legal Entity Name",
                  each Text.Trim(_, {" ", """", "'", "“", "”", "„", "‘", "’", "`", "´", "«", "»"}), type text}}),

    WithCC    = Table.AddColumn(Names, "Country Code", each Text.Start([Legal Entity Code], 2), type text),

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
    CEE       = {"BG","CZ","HR","HU","PL","RO","CS","RS","SI","SK","ME"},
    EAkeys    = {"europ assistance", "europäische reise", "európai utazási"},

    WithCtry  = Table.AddColumn(WithCC, "Country", each Record.FieldOrDefault(Countries, [Country Code], [Country Code]), type text),
    WithCcy   = Table.AddColumn(WithCtry, "Currency", each Record.FieldOrDefault(Ccy, [Country Code], "EUR"), type text),
    WithReg   = Table.AddColumn(WithCcy, "Regime", each
                  if List.Contains(SII, [Country Code]) then "Solvency II"
                  else if [Country Code] = "CH" then "SST (CH)"
                  else "Non-SII", type text),

    WithBU    = Table.AddColumn(WithReg, "BU", each
                  let n = Text.Lower([Legal Entity Name]), c = [Country Code] in
                  if List.AnyTrue(List.Transform(EAkeys, (k) => Text.Contains(n, k))) then "Europ Assistance"
                  else if c = "IT" then "Italy"
                  else if c = "DE" then "Germany"
                  else if List.Contains({"FR","LU","MC"}, c) then "France"
                  else if c = "AT" then "Austria"
                  else if List.Contains({"CH","LI"}, c) then "Switzerland"
                  else if List.Contains(CEE, c) or [Legal Entity Code] = "NL100" then "CEE"
                  else if List.Contains({"ES","PT","GR","IE","BE","NL","GB"}, c) then "International – Europe"
                  else if List.Contains({"AR","BR","CL","EC","MX","CO","PE","PA","GT"}, c) then "International – LatAm"
                  else "International – Other", type text),

    Final     = Table.AddColumn(WithBU, "BU Order", each
                  Record.FieldOrDefault(
                    [Italy = 1, Germany = 2, France = 3, Austria = 4, Switzerland = 5, CEE = 6,
                     #"Europ Assistance" = 7, #"International – Europe" = 8, #"International – LatAm" = 9],
                    [BU], 10), Int64.Type)
in
    Final
