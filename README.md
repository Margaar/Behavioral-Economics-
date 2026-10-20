>> 1) ПОРОГИ КАК МЕРЫ (по одной, Home > Nuova misura). Шаблон проверен твоей Thr Test.
>>    Положи их в папку "Thresholds" (Cartella di visualizzazione).

Thr SOL_RED = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SOL_RED" )

Thr SOL_AMBER = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SOL_AMBER" )

Thr SENS_RED = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SENS_RED" )

Thr SENS_AMBER = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SENS_AMBER" )

Thr REM_AMBER_DROP = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "REM_AMBER_DROP" )

Thr REM_RED_DROP = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "REM_RED_DROP" )

Thr CN_AMBER_INC = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CN_AMBER_INC" )

Thr CN_RED_INC = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CN_RED_INC" )

Thr RH_MIN_CASH = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "RH_MIN_CASH" )

Thr CHECK_TOL = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CHECK_TOL" )

Thr FTC_RED = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "FTC_RED" )

Thr SAA_AMBER_BREACH = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SAA_AMBER_BREACH" )

Thr SAA_RED_BREACH = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SAA_RED_BREACH" )

Thr SAA_SUM_TOL = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SAA_SUM_TOL" )

Thr LK_RED_DROP = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "LK_RED_DROP" )

Thr NL_COR_AMBER = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "NL_COR_AMBER" )

Thr NL_COR_RED = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "NL_COR_RED" )


>> 2) ЗАМЕНИ ЭТИ МЕРЫ ЦЕЛИКОМ (Ctrl+A в строке формул > вставить)

SOL Status =
VAR r   = [SOL Min Ratio]
VAR red = [Thr SOL_RED]
VAR amb = [Thr SOL_AMBER]
RETURN SWITCH ( TRUE (), ISBLANK ( r ), BLANK (), r < red, 3, r < amb, 2, 1 )

AoM Status =
VAR tol = [Thr CHECK_TOL]
VAR red = [Thr SOL_RED]
VAR amb = [Thr SOL_AMBER]
VAR r   = DIVIDE ( [AoM EOF EoP], [AoM SCR EoP] )
VAR ce  = ABS ( ROUND ( [AoM Check EOF], 0 ) )
VAR cs  = ABS ( ROUND ( [AoM Check SCR], 0 ) )
RETURN SWITCH ( TRUE (), ISBLANK ( r ), BLANK (), ce > tol || cs > tol, 3, r < red, 3, r < amb, 2, 1 )

AoM Headline =
VAR tol = [Thr CHECK_TOL]
VAR r0  = DIVIDE ( [AoM EOF BoP], [AoM SCR BoP] )
VAR r1  = DIVIDE ( [AoM EOF EoP], [AoM SCR EoP] )
VAR ok  = ABS ( ROUND ( [AoM Check EOF], 0 ) ) <= tol && ABS ( ROUND ( [AoM Check SCR], 0 ) ) <= tol
RETURN IF ( NOT ISBLANK ( r1 ),
    "Solvency ratio walk " & FORMAT ( r0, "0%" ) & " → " & FORMAT ( r1, "0%" )
    & " (" & FORMAT ( ( r1 - r0 ) * 100, "+0.0;-0.0;0.0" ) & " pp); "
    & IF ( ok, "EOF and SCR bridges reconcile.", "a bridge does NOT reconcile – check the page." ) )

SENS Status =
VAR a   = [SENS Worst After]
VAR red = [Thr SENS_RED]
VAR amb = [Thr SENS_AMBER]
RETURN SWITCH ( TRUE (), ISBLANK ( a ), BLANK (), a < red, 3, a < amb, 2, 1 )

FTC Status =
VAR tol = [Thr CHECK_TOL]
VAR lim = [Thr FTC_RED]
VAR y1  = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Free Tangible Capital" )
VAR f   = CALCULATE ( [FTC FTC], FactData[Year] = y1 )
VAR chk = ABS ( ROUND ( CALCULATE ( [FTC Check], FactData[Year] = y1 ), 0 ) )
RETURN SWITCH ( TRUE (), ISBLANK ( f ), BLANK (), chk > tol, 3, f < lim, 3, 1 )

FTC Headline =
VAR tol = [Thr CHECK_TOL]
VAR y0  = CALCULATE ( MIN ( FactData[Year] ), FactData[Topic] = "Free Tangible Capital" )
VAR y1  = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Free Tangible Capital" )
VAR f0  = CALCULATE ( [FTC FTC], FactData[Year] = y0 )
VAR f1  = CALCULATE ( [FTC FTC], FactData[Year] = y1 )
VAR chk = ROUND ( CALCULATE ( [FTC Check], FactData[Year] = y1 ), 0 )
RETURN IF ( NOT ISBLANK ( f1 ),
    "Free Tangible Capital " & FORMAT ( f0, "#,##0" ) & " in " & y0 & ", " & FORMAT ( f1, "#,##0" ) & " in " & y1
    & IF ( ABS ( chk ) <= tol, "; reconciliation closes.", "; reconciliation gap " & FORMAT ( chk, "+#,##0;-#,##0" ) & "." ) )

CN Status =
VAR d   = [CN Delta %]
VAR amb = [Thr CN_AMBER_INC]
VAR red = [Thr CN_RED_INC]
VAR tot = CALCULATE ( [CN Total Cur], FactData[Topic] = "Capital Needs" )
RETURN SWITCH ( TRUE (), ISBLANK ( tot ), BLANK (), ISBLANK ( d ), 1, d > red, 3, d > amb, 2, 1 )

TR Status =
VAR d   = [TR Delta %]
VAR amb = [Thr REM_AMBER_DROP]
VAR red = [Thr REM_RED_DROP]
RETURN SWITCH ( TRUE (), ISBLANK ( d ), BLANK (), d < -red, 3, d < -amb, 2, 1 )

RBC Status =
VAR d   = [RBC Delta %]
VAR amb = [Thr REM_AMBER_DROP]
VAR red = [Thr REM_RED_DROP]
RETURN SWITCH ( TRUE (), ISBLANK ( d ), BLANK (), d < -red, 3, d < -amb, 2, 1 )

RH Status =
VAR c   = [RH Cash End]
VAR lim = [Thr RH_MIN_CASH]
RETURN SWITCH ( TRUE (), ISBLANK ( c ), BLANK (), c < lim, 3, 1 )

SAA Sum Off =
VAR tol = [Thr SAA_SUM_TOL]
VAR t = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Entity], FactData[SAA Part] ),
            FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation" )
VAR x = ADDCOLUMNS ( t, "@s", CALCULATE ( SUM ( FactData[Value] ),
            FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
            FactData[SAA Col] = "Target %", FactData[SAA Order] <= 9 ) )
RETURN COUNTROWS ( FILTER ( x, NOT ISBLANK ( [@s] ) && [@s] <> 0 && ABS ( [@s] - 1 ) > tol ) ) + 0

SAA Status =
VAR tg  = [SAA Tgt MV]
VAR nB  = [SAA Breaches]
VAR amb = [Thr SAA_AMBER_BREACH]
VAR red = [Thr SAA_RED_BREACH]
RETURN SWITCH ( TRUE (), ISBLANK ( tg ), BLANK (), nB >= red, 3, nB >= amb || [SAA Sum Off] > 0, 2, 1 )

LK Status =
VAR a   = [LK VNP First]
VAR z   = [LK VNP Last]
VAR red = [Thr LK_RED_DROP]
VAR d   = DIVIDE ( z - a, ABS ( a ) )
RETURN SWITCH ( TRUE (), ISBLANK ( z ), BLANK (), ISBLANK ( d ), 1, d < -red, 3, d < 0, 2, 1 )

NL Status =
VAR c   = [NL CoR Last]
VAR amb = [Thr NL_COR_AMBER]
VAR red = [Thr NL_COR_RED]
RETURN SWITCH ( TRUE (), ISBLANK ( c ) || c = 0, BLANK (), c > red, 3, c > amb, 2, 1 )
