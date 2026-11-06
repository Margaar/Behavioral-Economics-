            IsNumTxt = (t) => t <> null and Text.Select(t, {"0".."9", ",", "."}) = t,
            BlockNm = if HLabel <> null and HLabel <> "" and not IsNumTxt(HLabel) and not IsPeriod(HLabel) then HLabel
                      else if not List.IsEmpty(List.Select(HLeft, (t) => not IsNumTxt(t)
                              and not Text.StartsWith(Text.Upper(t), "COMPANY #")))
                           then List.Select(HLeft, (t) => not IsNumTxt(t)
                              and not Text.StartsWith(Text.Upper(t), "COMPANY #")){0}
