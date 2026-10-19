=====================================================================
 FREE TANGIBLE CAPITAL — сборка страницы с нуля
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> КАК УСТРОЕН ЛИСТ (по фото Excel), у каждой компании 2 таблицы:
>>
>> 1) RECONCILIATION (Act 2025 ... Plan 2029): от капитала к FTC
>>    Local GAAP Shareholders' Equity
>>    + (-) Goodwill, (-) Revaluation, (-) Other reserves, (-) Foreseeable Div., (+) Subordinated
>>    = LGAAP Tangible Capital (TCAP)   [o/w Not Distributable / o/w Distributable]
>>    + LGAAP TCAP to cover Required Capital Buffer
>>    = LGAAP TCAP after Required Buffer coverage
>>    + (-) Not Distributable TCAP, (-) Distributable TCAP backed by Part. ...
>>    = Tangible Capital
>>    + (-) Tangible Capital to cover OTR - Upper Corridor
>>    = FREE TANGIBLE CAPITAL
>>    Проверка ES003 Plan 2027: 1482 - 316 + 40 - 384 - 216 - 343 - 101 = 162 ✓
>>
>> 2) YoY MOVEMENT (For 2026 ... Plan 2029): FTC (BoP) + движения = FTC (EoP)
>>    ES003 Plan 2027: 0 + 151 - 384 + 54 - 168 + 51 + 559 - 101 = 162 ✓


>> ШАГ 0. ПРОВЕРКА (3 минуты)
>> Dati > FactData > фильтр Topic = Free Tangible Capital. Посмотри:
>>   Block:  два названия, одно содержит "Reconciliation", другое "Movement".
>>   Entity: названия компаний (не пусто).
>>   Item:   строки обеих таблиц, в т.ч. "Free Tangible Capital (BoP)".
>>   SignHint: "-" / "+" у строк (-) Goodwill..., (+) Subordinated...
>> Если таблицы Movement нет или Block называется иначе — пришли фото
>> (тогда поменяем одно слово в колонке FTC Table).


>> ШАГ 1. ЧЕТЫРЕ КОЛОНКИ в FactData (Dati > FactData > Nuova colonna), по порядку

FTC Table =
IF ( FactData[Topic] = "Free Tangible Capital",
     IF ( CONTAINSSTRING ( FactData[Block], "Movement" ), "Movement", "Reconciliation" ) )

FTC RelRow =
IF ( FactData[Topic] = "Free Tangible Capital",
     FactData[RowNo]
     - CALCULATE ( MIN ( FactData[RowNo] ), ALLEXCEPT ( FactData, FactData[Topic], FactData[BlockNo] ) ) )

>> Название строки; строки таблицы Movement получают невидимый символ,
>> чтобы не смешиваться с одноимёнными строками Reconciliation
FTC Line =
IF ( FactData[Topic] = "Free Tangible Capital",
     FactData[Item] & IF ( FactData[FTC Table] = "Movement", UNICHAR ( 8203 ), "" ) )

FTC Order =
IF ( FactData[Topic] = "Free Tangible Capital",
     CALCULATE ( MIN ( FactData[FTC RelRow] ),
         ALLEXCEPT ( FactData, FactData[Topic], FactData[Block], FactData[Item] ) ) )

>> Потом: колонка FTC Line > Ordina per colonna > FTC Order.


>> ШАГ 2. МЕРЫ (Home > Nuova misura), по одной

FTC Equity =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Reconciliation",
    CONTAINSSTRING ( FactData[Item], "Shareholders" ) )

FTC TCAP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Reconciliation",
    CONTAINSSTRING ( FactData[Item], "LGAAP Tangible Capital" ) )

FTC Buffer =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Reconciliation",
    CONTAINSSTRING ( FactData[Item], "to cover Required Capital Buffer" ) )

FTC TC =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Reconciliation",
    FactData[Item] = "Tangible Capital" )

FTC OTR =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Reconciliation",
    CONTAINSSTRING ( FactData[Item], "to cover OTR" ) )

FTC FTC =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Reconciliation",
    FactData[Item] = "Free Tangible Capital" )

>> Проверка 1: Equity + все строки со знаком (-)/(+) + Buffer = FTC  (должно быть 0)
FTC Check =
[FTC Equity]
+ CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
      REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
      FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Reconciliation",
      NOT ISBLANK ( FactData[SignHint] ) )
+ [FTC Buffer]
- [FTC FTC]

>> Проверка 2: BoP + движения = EoP  (должно быть 0)
FTC Move Check =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Movement",
    NOT CONTAINSSTRING ( FactData[Item], "(EoP)" ) )
- CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
    FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Movement",
    CONTAINSSTRING ( FactData[Item], "(EoP)" ) )

>> Для водопада "от капитала к FTC": только стартовая строка и строки-корректировки
FTC Bridge =
VAR it  = SELECTEDVALUE ( FactData[Item] )
VAR tbl = SELECTEDVALUE ( FactData[FTC Table] )
VAR sg  = SELECTEDVALUE ( FactData[SignHint] )
RETURN
    IF ( tbl = "Reconciliation"
             && ( CONTAINSSTRING ( it, "Shareholders" ) || NOT ISBLANK ( sg )
                  || CONTAINSSTRING ( it, "to cover Required Capital Buffer" ) ),
         [Value Current] )

>> Для водопада "движение FTC за год": BoP и движения, без EoP
FTC Move =
VAR it  = SELECTEDVALUE ( FactData[Item] )
VAR tbl = SELECTEDVALUE ( FactData[FTC Table] )
RETURN IF ( tbl = "Movement" && NOT CONTAINSSTRING ( it, "(EoP)" ), [Value Current] )

>> Цвет фона итоговых строк в матрицах (как в Excel)
FTC Row Color =
VAR it = SELECTEDVALUE ( FactData[Item] )
RETURN
    SWITCH ( TRUE (),
        CONTAINSSTRING ( it, "Free Tangible Capital" ), "#F4D3B5",
        it = "Tangible Capital"
            || CONTAINSSTRING ( it, "Shareholders" )
            || CONTAINSSTRING ( it, "LGAAP Tangible Capital" )
            || CONTAINSSTRING ( it, "after Required Buffer" ), "#E3E3E3",
        BLANK () )


>> ШАГ 3. КОММЕНТАРИЙ
>> Без выбора: последний год плана. Срез/клик по ячейке: этот год. Клик по строке: строка.

FTC Comment =
VAR fmt  = "#,##0"
VAR fmtD = "+#,##0;-#,##0;0"
VAR nl   = UNICHAR ( 10 )
VAR yrs  = CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( DimItem ),
               REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
               FactData[Topic] = "Free Tangible Capital", NOT ISBLANK ( FactData[Value] ) )
VAR y0   = MINX ( yrs, FactData[Year] )
VAR y1   = MAXX ( yrs, FactData[Year] )
VAR period = IF ( y0 = y1, "in " & y1, y0 & "-" & y1 )
VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Free Tangible Capital" )
VAR who  = IF ( nEnt = 1,
               CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Free Tangible Capital" ),
               nEnt & " companies" )
VAR ftc1 = CALCULATE ( [FTC FTC], FactData[Year] = y1 )
VAR ftc0 = CALCULATE ( [FTC FTC], FactData[Year] = y0 )
VAR eq1  = CALCULATE ( [FTC Equity], FactData[Year] = y1 )
VAR tcp1 = CALCULATE ( [FTC TCAP], FactData[Year] = y1 )
VAR buf1 = CALCULATE ( [FTC Buffer], FactData[Year] = y1 )
VAR tc1  = CALCULATE ( [FTC TC], FactData[Year] = y1 )
VAR otr1 = CALCULATE ( [FTC OTR], FactData[Year] = y1 )
VAR conv = DIVIDE ( ftc1, tcp1 )
VAR moves =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( FactData[FTC Line] ), "@v", [Value Current] ),
        REMOVEFILTERS ( DimItem ),
        REMOVEFILTERS ( FactData[FTC Line], FactData[FTC Order], FactData[FTC Table] ),
        FactData[Topic] = "Free Tangible Capital", FactData[FTC Table] = "Movement",
        NOT CONTAINSSTRING ( FactData[Item], "(BoP)" ), NOT CONTAINSSTRING ( FactData[Item], "(EoP)" ) )
VAR top3Txt =
    CONCATENATEX ( TOPN ( 3, FILTER ( moves, [@v] <> 0 ), ABS ( [@v] ), DESC ),
        SUBSTITUTE ( FactData[FTC Line], UNICHAR ( 8203 ), "" ) & " (" & FORMAT ( [@v], fmtD ) & ")",
        "; ", ABS ( [@v] ), DESC )
VAR chk  = ROUND ( CALCULATE ( [FTC Check], FactData[Year] = y1 ), 0 )
VAR chkM = ROUND ( CALCULATE ( [FTC Move Check], FactData[Year] = y1 ), 0 )
VAR pageText =
    "FREE TANGIBLE CAPITAL: " & UPPER ( who ) & ", " & period & nl & nl
    & IF ( ftc1 = 0 || ISBLANK ( ftc1 ),
           "No Free Tangible Capital in " & y1 & ": the required capital buffer and the OTR coverage absorb the whole tangible capital.",
           "Free Tangible Capital " & y1 & ": " & FORMAT ( ftc1, fmt )
           & IF ( y0 <> y1, " (" & FORMAT ( ftc1 - ftc0, fmtD ) & " vs " & y0 & ")", "" ) & "." )
    & nl
    & "Bridge " & y1 & ": shareholders' equity " & FORMAT ( eq1 + 0, fmt )
    & " → LGAAP tangible capital " & FORMAT ( tcp1 + 0, fmt )
    & "; required buffer " & FORMAT ( buf1 + 0, fmt )
    & " → tangible capital " & FORMAT ( tc1 + 0, fmt )
    & "; OTR upper corridor " & FORMAT ( otr1 + 0, fmt )
    & " → FTC " & FORMAT ( ftc1 + 0, fmt ) & "."
    & IF ( tcp1 > 0, nl & FORMAT ( conv, "0%" ) & " of LGAAP tangible capital is free.", "" )
    & IF ( top3Txt <> "", nl & "Main drivers of the FTC movement " & period & ": " & top3Txt & ".", "" )
    & nl & nl
    & IF ( chk = 0 && chkM = 0, "Check: reconciliation and YoY movement are consistent.",
           "Check: differences found (reconciliation " & FORMAT ( chk, fmtD ) & ", movement " & FORMAT ( chkM, fmtD )
           & ") in " & y1 & ". Please review the input file." )
VAR ln  = SELECTEDVALUE ( FactData[FTC Line] )
VAR byYear =
    CONCATENATEX ( yrs,
        FactData[Year] & ": " & FORMAT ( CALCULATE ( [Value Current], FactData[Topic] = "Free Tangible Capital" ) + 0, fmt ),
        " | ", FactData[Year], ASC )
VAR rowText =
    UPPER ( SUBSTITUTE ( ln, UNICHAR ( 8203 ), "" ) ) & ": " & who & nl & nl
    & "By year: " & byYear & "."
    & IF ( ftc1 <> 0 && NOT CONTAINSSTRING ( ln, "Free Tangible Capital" ),
           nl & "In " & y1 & " this line equals "
           & FORMAT ( DIVIDE ( CALCULATE ( [Value Current], FactData[Topic] = "Free Tangible Capital", FactData[Year] = y1 ), ftc1 ), "0%" )
           & " of Free Tangible Capital.", "" )
RETURN IF ( HASONEVALUE ( FactData[FTC Line] ), rowText, pageText )


>> ШАГ 4. СТРАНИЦА "Free Tangible Capital"
>>
>> Фильтр страницы: Topic = Free Tangible Capital; PeriodType = Actual, Forecast, Plan.
>> Срезы: Entity (синхронизирован), Scale = Millions. RefScenario/CompScenario не нужны.
>>
>> a) МАТРИЦА 1 "Free Tangible Capital Reconciliation vs Local GAAP Shareholders' Equity"
>>    Righe: FTC Line   Colonne: DimPeriod[Period]   Valori: Value Current
>>    Фильтр визуала: FTC Table = Reconciliation
>>    Totali di riga / colonna: Disattivato
>>    Фон итогов: Formattazione condizionale > Colore sfondo > Stile "Valore campo" > FTC Row Color
>>
>> b) МАТРИЦА 2 "Free Tangible Capital YoY Movement"
>>    Копия матрицы 1, фильтр визуала FTC Table = Movement
>>
>>    Если выбрано несколько компаний, поставь в Righe сначала Entity, потом FTC Line —
>>    будет как в Excel: компания, под ней две таблицы.
>>
>> c) ВОДОПАД "From equity to Free Tangible Capital" (один год)
>>    Categoria: FTC Line   Asse Y: FTC Bridge
>>    Фильтр визуала: FTC Table = Reconciliation
>>    Ordina asse > FTC Line > Crescente
>>    Срез Year (Selezione singola) рядом; Modifica interazioni: влияет только на водопады.
>>    Первый столбец = Equity, последний "Totale" = FTC.
>>
>> d) ВОДОПАД "FTC movement" (один год)
>>    Categoria: FTC Line   Asse Y: FTC Move
>>    Фильтр визуала: FTC Table = Movement
>>    Первый столбец = FTC (BoP), "Totale" = FTC (EoP).
>>
>> e) СТОЛБЦЫ "FTC by company"
>>    Istogramma in pila: Asse X = Period (PeriodType = Plan), Asse Y = FTC FTC, Legenda = Entity
>>    Срез Entity > Modifica interazioni > над этим графиком "Nessuno" (видно все компании).
>>
>> f) КОММЕНТАРИЙ: таблица с FTC Comment, A capo automatico = Attivato.
>>    Карточки (по желанию): FTC FTC, FTC Check, FTC Move Check.


>> ПРОВЕРКА (Generali España ES003, Scale = Millions)
>>   Shareholders' Equity: Act 2025 = 688, Plan 2027 = 1.482
>>   LGAAP TCAP:           Act 2025 = 483, Plan 2027 = 822
>>   Free Tangible Capital: Act 2025 = 0, Plan 2027 = 162, Plan 2028 = 214, Plan 2029 = 125
>>   Водопад Plan 2027: старт 1.482 → Totale 162
>>   FTC Check и FTC Move Check = 0 (± округление)
