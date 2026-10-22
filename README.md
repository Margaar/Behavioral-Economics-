    All     = if IncludeCMP then Table.Combine({WithSrc, CmpRows}) else WithSrc,



        Codes   = List.Buffer(List.RemoveNulls(CompanyMap[Code])),
    InScope = if List.IsEmpty(Codes) then Per
              else Table.SelectRows(Per, each List.Contains(Codes, Text.From([LECode]))),


              
