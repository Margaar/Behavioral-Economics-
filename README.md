
    EntOf = (h) =>
        let p0 = P0Of(h),
            Rws    = List.Select(List.Numbers(h, 7, -1), (r) => r >= 0),
            IsComp = (t) => t <> null and Text.StartsWith(Text.Upper(t), "COMPANY #"),
            IsName = (t) => t <> null and List.AnyTrue(List.Transform(EntNames, (n) => Text.StartsWith(Text.Upper(t), n))),
            Find   = (cols as list, f) => List.First(List.Select(List.Combine(List.Transform(Rws,
                         (r) => List.Transform(cols, (c) => Txt(Rows{r}{c})))), f), null),
            Lc = List.Reverse(List.Select(WinCols, (c) => c < p0)),
            Oc = List.Select(WinCols, (c) => c >= p0),
            Nc = List.Reverse(List.Select(List.Numbers(C0 - 3, 3), (c) => c >= 0)),
            Hits = { Find(Lc, IsComp), Find(Oc, IsComp), Find(Nc, IsComp),
                     Find(Lc, IsName), Find(Oc, IsName), Find(Nc, IsName) }
        in List.First(List.RemoveNulls(Hits), null),

        





VAR it = SELECTEDVALUE ( FactData[LK Line] )
RETURN IF ( CONTAINSSTRING ( it, "(%)" ) || CONTAINSSTRING ( it, "Guarantee-linked" )
            || CONTAINSSTRING ( it, "Average Guarantee" ), "0.0%", "#,##0" )
        
