let
    T = Table.Distinct(Table.SelectColumns(FactData, {"Topic", "Item"})),
    S = Table.Sort(T, {{"Topic", Order.Ascending}, {"Item", Order.Ascending}})
in
    S
