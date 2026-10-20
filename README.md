    EntOf = (h) =>
        let p0 = P0Of(h),
            Rws    = List.Select(List.Numbers(h, 7, -1), (r) => r >= 0),
            IsComp = (t) => t <> null and Text.StartsWith(Text.Upper(t), "COMPANY #"),
            IsName = (t) => t <> null and List.AnyTrue(List.Transform(EntNames, (n) => Text.StartsWith(Text.Upper(t), n))),
            Find   = (cols as list, f) => List.First(List.Select(List.Combine(List.Transform(Rws,
                         (r) => List.Transform(cols, (c) => Txt(Rows{r}{c})))), f), null),
            Lc = List.Reverse(List.Select(WinCols, (c) => c < p0)),
            Oc = List.Select(WinCols, (c) => c >= p0),
            Ac = List.Numbers(0, NC),
            Hits = { Find(Lc, IsComp), Find(Oc, IsComp), Find(Ac, IsComp),
                     Find(Lc, IsName), Find(Oc, IsName), Find(Ac, IsName) }
        in List.First(List.RemoveNulls(Hits), null),







VAR it = SELECTEDVALUE ( FactData[LK Line] )
RETURN IF ( CONTAINSSTRING ( it, "(%)" ) || CONTAINSSTRING ( it, "Guarantee-linked" )
            || CONTAINSSTRING ( it, "Average Guarantee" ), "0.0%", "#,##0" )
        
