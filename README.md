FTC Comment HTML =
VAR _tbl =
    FILTER (
        ADDCOLUMNS (
            SUMMARIZE ( EntityMap, EntityMap[Legal Entity Code], EntityMap[Legal Entity Name] ),
            "@Pre",  [FTC Pre],
            "@Post", [FTC Post],
            "@D",    [FTC Δ]
        ),
        NOT ISBLANK ( [@Post] ) && [@Post] <> 0
    )
VAR _n    = COUNTROWS ( _tbl )
VAR _pre  = SUMX ( _tbl, [@Pre] )
VAR _post = SUMX ( _tbl, [@Post] )
VAR _d    = _post - _pre
VAR _dp   = DIVIDE ( _d, ABS ( _pre ) )
VAR _up   = COUNTROWS ( FILTER ( _tbl, [@D] > 0 ) )
VAR _dn   = COUNTROWS ( FILTER ( _tbl, [@D] < 0 ) )
VAR _col  = IF ( _d >= 0, "#1E7B34", "#C00000" )
VAR _top  = TOPN ( 3, FILTER ( _tbl, [@D] > 0 ), [@D], DESC )
VAR _bot  = TOPN ( 3, FILTER ( _tbl, [@D] < 0 ), [@D], ASC )
VAR _li   = "<li style='margin:2px 0'>"
VAR _topL =
    CONCATENATEX ( _top,
        _li & "<b>" & EntityMap[Legal Entity Name] & "</b> <span style='color:#1E7B34'>"
            & FORMAT ( [@D], "+#,##0.0;-#,##0.0" ) & " €m</span></li>", "", [@D], DESC )
VAR _botL =
    CONCATENATEX ( _bot,
        _li & "<b>" & EntityMap[Legal Entity Name] & "</b> <span style='color:#C00000'>"
            & FORMAT ( [@D], "+#,##0.0;-#,##0.0" ) & " €m</span></li>", "", [@D], ASC )
RETURN
"<div style='font-family:Segoe UI;font-size:12px;color:#333;padding:12px 14px;
      border-left:5px solid #C5281C;background:#FAFAFA;border-radius:4px'>"
& "<div style='font-size:14px;font-weight:600;color:#C5281C;margin-bottom:6px'>
      FTC YE26 – Solvency II Review impact</div>"
& "<div>Perimeter: <b>" & _n & "</b> SII companies</div>"
& "<div style='margin:6px 0;font-size:16px'>"
&     FORMAT ( _pre, "#,##0.0" ) & " → <b>" & FORMAT ( _post, "#,##0.0" ) & " €m</b> "
&     "<span style='color:" & _col & ";font-weight:600'>("
&     FORMAT ( _d, "+#,##0.0;-#,##0.0" ) & " €m, " & FORMAT ( _dp, "+0.0%;-0.0%" ) & ")</span></div>"
& "<div style='color:#666'>" & _up & " companies up · " & _dn & " down</div>"
& IF ( NOT ISEMPTY ( _top ),
      "<div style='margin-top:8px;font-weight:600'>Top increases</div><ul style='margin:2px 0 0 16px;padding:0'>" & _topL & "</ul>" )
& IF ( NOT ISEMPTY ( _bot ),
      "<div style='margin-top:6px;font-weight:600'>Top decreases</div><ul style='margin:2px 0 0 16px;padding:0'>" & _botL & "</ul>" )
& "</div>"
