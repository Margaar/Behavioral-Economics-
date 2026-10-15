=====================================================================
 ИНСТРУКЦИЯ: HMS IMPACTS, SOLVENCY AoM, SENSITIVITIES, FTC
 (Power BI, шаблон Capital Deep Dive)
=====================================================================

КАК ЧИТАТЬ ЭТОТ ФАЙЛ
- Строки с ">>" — это объяснения. Их НЕ копировать.
- Копировать только код: от строки "Имя меры =" до пустой строки.
- Внутри кода комментариев нет специально: при копировании
  комментарии "--" иногда ломают меру (как было со 'status').
- Меры: Home > Nuova misura, по одной, сверху вниз.
- Колонки: вид Dati > нужная таблица > Nuova colonna.
- После каждого блока: Ctrl+S.

ПОРЯДОК РАБОТЫ (рекомендую)
  1. Solvency AoM   (самая важная, движение Solvency Ratio)
  2. Sensitivities  (стресс-тесты)
  3. FTC
  4. HMS Impacts
  На каждую страницу ~1-1,5 часа.


=====================================================================
 0. ОБЩИЕ ПРАВИЛА ДЛЯ КАЖДОЙ СТРАНИЦЫ
=====================================================================

ШАГ 0 — ВСЕГДА СНАЧАЛА ПРОВЕРИТЬ СТРУКТУРУ (5 минут)
  Вид Dati > FactData > фильтр Topic = <тема страницы>.
  Посмотреть колонки: Block, Entity, Section, Item, Period, Measure.
  >> Если Entity заполнен  -> на странице нужен срез Entity.
  >> Если Measure НЕ пустой (например Gross / Net) -> в матрице
     поставить Measure в Colonne рядом с Year ИЛИ отфильтровать
     одно значение. Иначе Power BI сложит Gross + Net.
  >> Если структура не похожа на то, что описано ниже —
     пришли мне фото, поправлю меры.

КАРКАС СТРАНИЦЫ
  - Формат 16:9, заголовок, кнопки навигации (как на других страницах).
  - Фильтр страницы: Topic = <тема>.
  - Срез Scale справа сверху.
  - Срез Entity — только если в данных есть компании.
    Visualizza > Sincronizza filtri dei dati: отметить эту страницу.
  - Таблица комментария справа или снизу:
    Tabella > Valori = мера комментария,
    Formato > Valori > A capo automatico = Attivato,
    Modifica interazioni: матрица и графики -> комментарий = Filtro.

ТЕМЫ (Topic) — писать точно так:
  "Solvency AoM"   "Sensitivities"   "Free Tangible Capital"   "HMS Impacts"


=====================================================================
 1. SOLVENCY AoM  (Analysis of Movement: как менялся Solvency Ratio)
=====================================================================

>> Что показываем: как Solvency Ratio дошёл от начала года (BoP)
>> до конца года (EoP). Каждая строка EOF и SCR переводится
>> в "пункты" коэффициента (pp). Это стандартный слайд
>> "Solvency Ratio walk". Сумма всех пунктов = EoP - BoP.
>> Как считается вклад:
>>   строки EOF: движение / SCR на начало
>>   строки SCR: доля движения SCR * (EOF_конец/SCR_конец - EOF_конец/SCR_начало)
>> Так шаги точно складываются в итог.

--- 1.1 Три вычисляемые колонки в FactData ---
>> AoM Line: название строки с префиксом EOF/SCR, потому что
>> одинаковые названия (например "Market movements") могут быть
>> и в EOF, и в SCR.

AoM Line =
IF ( FactData[Topic] = "Solvency AoM",
     IF ( LEFT ( FactData[Section], 3 ) = "SCR", "SCR | ", "EOF | " ) & FactData[Item] )

>> AoM RelRow: номер строки внутри блока компании (для порядка).

AoM RelRow =
IF ( FactData[Topic] = "Solvency AoM",
     FactData[RowNo]
     - CALCULATE ( MIN ( FactData[RowNo] ), ALLEXCEPT ( FactData, FactData[Topic], FactData[BlockNo] ) ) )

>> AoM Order: один номер на каждую строку AoM (для сортировки).

AoM Order =
IF ( FactData[Topic] = "Solvency AoM",
     CALCULATE ( MIN ( FactData[AoM RelRow] ),
         ALLEXCEPT ( FactData, FactData[Topic], FactData[Section], FactData[Item] ) ) )

>> Затем: выделить колонку AoM Line > Ordina per colonna > AoM Order.
>> Если Power BI выдаст ошибку сортировки — пришли фото.

--- 1.2 Меры: начало и конец EOF / SCR ---

AoM EOF BoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "EOF", CONTAINSSTRING ( FactData[Item], "(BoP)" ) )

AoM EOF EoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "EOF", CONTAINSSTRING ( FactData[Item], "(EoP)" ) )

AoM SCR BoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "SCR", CONTAINSSTRING ( FactData[Item], "(BoP)" ) )

AoM SCR EoP =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order] ),
    FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "SCR", CONTAINSSTRING ( FactData[Item], "(EoP)" ) )

AoM Ratio BoP = DIVIDE ( [AoM EOF BoP], [AoM SCR BoP] )

AoM Ratio EoP = DIVIDE ( [AoM EOF EoP], [AoM SCR EoP] )

--- 1.3 Мера: вклад строки в коэффициент (в пунктах) ---
>> Для строки EOF (BoP) мера возвращает начальный коэффициент,
>> чтобы водопад начинался с него. Строки EoP пустые:
>> водопад сам рисует итоговый столбец = EoP ratio.

AoM Ratio Walk =
VAR ln  = SELECTEDVALUE ( FactData[AoM Line] )
VAR e0  = [AoM EOF BoP]
VAR e1  = [AoM EOF EoP]
VAR s0  = [AoM SCR BoP]
VAR s1  = [AoM SCR EoP]
VAR v   = [Value Current]
VAR isBoP = CONTAINSSTRING ( ln, "(BoP)" )
VAR isEoP = CONTAINSSTRING ( ln, "(EoP)" )
RETURN
    SWITCH ( TRUE (),
        NOT HASONEVALUE ( FactData[AoM Line] ), BLANK (),
        isEoP, BLANK (),
        isBoP && LEFT ( ln, 3 ) = "EOF", DIVIDE ( e0, s0 ) * 100,
        isBoP, BLANK (),
        LEFT ( ln, 3 ) = "EOF", DIVIDE ( v, s0 ) * 100,
        LEFT ( ln, 3 ) = "SCR", DIVIDE ( v, s1 - s0 ) * ( DIVIDE ( e1, s1 ) - DIVIDE ( e1, s0 ) ) * 100,
        BLANK () )

--- 1.4 Проверки (контроль качества данных) ---
>> Должны быть 0. Если нет — в шаблоне есть строка-подитог,
>> которая считается дважды, или ошибка в файле.

AoM Check EOF =
[AoM EOF BoP]
+ CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order] ),
      FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "EOF",
      NOT CONTAINSSTRING ( FactData[Item], "(BoP)" ), NOT CONTAINSSTRING ( FactData[Item], "(EoP)" ) )
- [AoM EOF EoP]

AoM Check SCR =
[AoM SCR BoP]
+ CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order] ),
      FactData[Topic] = "Solvency AoM", LEFT ( FactData[Section], 3 ) = "SCR",
      NOT CONTAINSSTRING ( FactData[Item], "(BoP)" ), NOT CONTAINSSTRING ( FactData[Item], "(EoP)" ) )
- [AoM SCR EoP]

--- 1.5 Визуалы на странице ---
  a) Срез Year: Selezione singola = Attivato (AoM всегда за один год).
  b) Срез Entity (одна компания).
  c) Карточки: AoM Ratio BoP, AoM Ratio EoP (формат %), AoM Check EOF, AoM Check SCR.
  d) ВОДОПАД "Solvency Ratio walk" (главный визуал):
       Categoria = FactData[AoM Line]
       Asse Y    = AoM Ratio Walk
       Ordina asse > AoM Line > Crescente
       Etichette dati: Attivato, формат 0.0
       >> Первый столбец = коэффициент на начало, последний "Totale" = на конец.
  e) Два маленьких водопада в деньгах (EOF и SCR):
       Categoria = AoM Line, Asse Y = Value Current
       Фильтр визуала AoM Line: Filtro avanzato
         "inizia con" EOF   E   "non contiene" (EoP)      (для второго: SCR)
  f) Матрица: Righe = AoM Line, Valori = Value Current, AoM Ratio Walk.

--- 1.6 Комментарий ---

AoM Comment =
VAR hardLim = 1.10
VAR softLim = 1.30
VAR fmt  = "#,##0"
VAR fmtP = "+0.0;-0.0;0.0"
VAR nl   = UNICHAR ( 10 )
VAR y    = MAX ( FactData[Year] )
VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency AoM" )
VAR who  = IF ( nEnt = 1,
               CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Solvency AoM" ),
               nEnt & " companies (simple sum)" )
VAR r0 = [AoM Ratio BoP]
VAR r1 = [AoM Ratio EoP]
VAR rSol = [SOL Ratio]
VAR st =
    IF ( ISBLANK ( r1 ), "cannot be calculated",
    IF ( r1 < hardLim, "BELOW the " & FORMAT ( hardLim, "0%" ) & " threshold: action needed",
    IF ( r1 < softLim, "between the " & FORMAT ( hardLim, "0%" ) & " and " & FORMAT ( softLim, "0%" ) & " thresholds",
         "above the " & FORMAT ( softLim, "0%" ) & " threshold" ) ) )
VAR lines =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( FactData[AoM Line] ), "@pp", [AoM Ratio Walk], "@v", [Value Current] ),
        REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[AoM Line], FactData[AoM Order] ),
        FactData[Topic] = "Solvency AoM" )
VAR movers = FILTER ( lines, NOT CONTAINSSTRING ( FactData[AoM Line], "(BoP)" )
                             && NOT CONTAINSSTRING ( FactData[AoM Line], "(EoP)" ) && [@pp] <> 0 )
VAR eofPP = SUMX ( FILTER ( movers, LEFT ( FactData[AoM Line], 3 ) = "EOF" ), [@pp] )
VAR scrPP = SUMX ( FILTER ( movers, LEFT ( FactData[AoM Line], 3 ) = "SCR" ), [@pp] )
VAR top3  = TOPN ( 3, movers, ABS ( [@pp] ), DESC )
VAR top3Txt = CONCATENATEX ( top3, FactData[AoM Line] & " (" & FORMAT ( [@pp], fmtP ) & " pp)", "; ", ABS ( [@pp] ), DESC )
VAR chkE = ROUND ( [AoM Check EOF], 0 )
VAR chkS = ROUND ( [AoM Check SCR], 0 )
VAR pageText =
    "SOLVENCY RATIO MOVEMENT " & y & ": " & UPPER ( who ) & nl & nl
    & "Solvency ratio moved from " & FORMAT ( r0, "0.0%" ) & " to " & FORMAT ( r1, "0.0%" )
    & " (" & FORMAT ( ( r1 - r0 ) * 100, fmtP ) & " pp), " & st & "." & nl
    & "Own funds movements: " & FORMAT ( eofPP, fmtP ) & " pp; SCR movements: " & FORMAT ( scrPP, fmtP ) & " pp." & nl
    & IF ( top3Txt <> "", "Main drivers: " & top3Txt & ".", "No movements reported." )
    & nl & nl
    & IF ( chkE = 0 && chkS = 0, "Check: EOF and SCR movements reconcile with the closing balances.",
           "Check: movements do not reconcile (EOF diff " & FORMAT ( chkE, fmt ) & ", SCR diff " & FORMAT ( chkS, fmt ) & "). Please review the input file." )
    & IF ( NOT ISBLANK ( rSol ) && ROUND ( ( r1 - rSol ) * 1000, 0 ) <> 0,
           nl & "Note: closing ratio differs from the Solvency sheet (" & FORMAT ( rSol, "0.0%" ) & ").", "" )
VAR ln  = SELECTEDVALUE ( FactData[AoM Line] )
VAR v   = [Value Current]
VAR pp  = [AoM Ratio Walk]
VAR secTot = IF ( LEFT ( ln, 3 ) = "EOF", eofPP, scrPP )
VAR rowText =
    UPPER ( ln ) & " " & y & ": " & who & nl & nl
    & "Amount: " & FORMAT ( v + 0, fmt ) & "." & nl
    & IF ( CONTAINSSTRING ( ln, "(BoP)" ) || CONTAINSSTRING ( ln, "(EoP)" ),
           "Opening/closing balance.",
           "Impact on the solvency ratio: " & FORMAT ( pp, fmtP ) & " pp"
           & IF ( secTot <> 0, " (" & FORMAT ( DIVIDE ( pp, secTot ), "0%" ) & " of all " & LEFT ( ln, 3 ) & " movements)", "" ) & "." )
RETURN
    IF ( NOT HASONEVALUE ( FactData[Year] ), "Select one year to see the solvency ratio movement.",
        IF ( HASONEVALUE ( FactData[AoM Line] ), rowText, pageText ) )


=====================================================================
 2. SENSITIVITIES  (лист "Sens on Net Result & Solvency")
=====================================================================

>> Что показываем: как каждый шок (+50 bps ставки, -25% акции,
>> спреды...) меняет Net Result, Own Funds и Solvency Ratio,
>> и какой коэффициент будет ПОСЛЕ шока по сравнению с порогом 110%.

>> ВАЖНО, ШАГ 0: в Dati, FactData, Topic = Sensitivities, посмотри,
>> где стоят названия шоков (в Item или в Section) и где написано
>> "Net Result" / "Own Funds" / "Solvency". Две колонки ниже
>> угадывают это автоматически. Проверь их глазами после создания.

--- 2.1 Две вычисляемые колонки в FactData ---

Sens Metric =
VAR ctx = UPPER ( FactData[Block] & " " & FactData[Section] & " " & FactData[Measure] & " " & FactData[Item] )
RETURN
    IF ( FactData[Topic] = "Sensitivities",
        SWITCH ( TRUE (),
            CONTAINSSTRING ( ctx, "SOLVENCY" ), "Solvency ratio (pp)",
            CONTAINSSTRING ( ctx, "OWN FUNDS" ), "Eligible Own Funds",
            CONTAINSSTRING ( ctx, "NET RESULT" ), "Net Result",
            "Other" ) )

Sens Shock =
IF ( FactData[Topic] = "Sensitivities",
     IF ( LEFT ( UPPER ( FactData[Item] ), 9 ) = "IMPACT ON", FactData[Section], FactData[Item] ) )

>> Проверка: в Dati отфильтруй Topic = Sensitivities. В Sens Shock
>> должны быть шоки (+50 bps..., -25% Equity...), в Sens Metric
>> — три значения. Если не так — фото мне.

--- 2.2 Меры ---
>> ppFactor: если в Dati влияние на коэффициент записано как -0,05
>> (= -5 пунктов) оставь 100. Если записано как -5 — поставь 1.
>> Проверяй со срезом Scale = Units.

Sens NR =
CALCULATE ( [Value Current], FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Net Result" )

Sens EOF =
CALCULATE ( [Value Current], FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Eligible Own Funds" )

Sens Ratio pp =
VAR ppFactor = 100
RETURN
    CALCULATE ( [Value Current], FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Solvency ratio (pp)" ) * ppFactor

Sens Year =
CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[Sens Shock], FactData[Sens Metric] ), REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Sensitivities", NOT ISBLANK ( FactData[Value] ) )

Sens Base Ratio =
VAR y = [Sens Year]
RETURN
    CALCULATE ( [SOL Ratio], REMOVEFILTERS ( FactData[Sens Shock], FactData[Sens Metric] ), FactData[Year] = y ) * 100

Sens Ratio After =
IF ( NOT ISBLANK ( [Sens Ratio pp] ), [Sens Base Ratio] + [Sens Ratio pp] )

--- 2.3 Визуалы ---
  a) Срез Entity (одна компания).
  b) Карточка: Sens Base Ratio (подпись "Solvency ratio before shocks, %").
  c) "Торнадо" по коэффициенту: Grafico a barre raggruppate
       Asse Y = FactData[Sens Shock]
       Asse X = Sens Ratio pp
       Ordina > Sens Ratio pp > Crescente (худший шок сверху)
       Colori > fx: правило "< 0 красный, >= 0 зелёный"
       Фильтр визуала: Sens Shock "non inizia con" [   и   "non inizia con" <
       (убирает пустые шаблоны [custom to specify], <Life Model Change>)
  d) Второе торнадо: то же, Asse X = Sens NR.
  e) Таблица: Sens Shock | Sens NR | Sens EOF | Sens Ratio pp | Sens Ratio After
       Formattazione condizionale на Sens Ratio After:
       < 110 красный фон, 110-130 жёлтый, > 130 зелёный.
       (Пороги одинаковые с Solvency: поправь, если у компании другие.)

--- 2.4 Комментарий ---

Sens Comment =
VAR hardLim = 110
VAR softLim = 130
VAR fmt  = "#,##0"
VAR fmtP = "+0.0;-0.0;0.0"
VAR nl   = UNICHAR ( 10 )
VAR y    = [Sens Year]
VAR base = [Sens Base Ratio]
VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Sensitivities" )
VAR who  = IF ( nEnt = 1,
               CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Sensitivities" ),
               nEnt & " companies (simple sum)" )
VAR shocks =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( FactData[Sens Shock] ), "@pp", [Sens Ratio pp], "@nr", [Sens NR] ),
        REMOVEFILTERS ( FactData[Sens Shock] ), REMOVEFILTERS ( DimItem ),
        FactData[Topic] = "Sensitivities", FactData[Year] = y )
VAR real = FILTER ( shocks, NOT ISBLANK ( FactData[Sens Shock] )
                            && NOT ( LEFT ( FactData[Sens Shock], 1 ) IN { "[", "<" } )
                            && ( [@pp] <> 0 || [@nr] <> 0 ) )
VAR worst   = TOPN ( 1, real, [@pp], ASC )
VAR wName   = MAXX ( worst, FactData[Sens Shock] )
VAR wPP     = MAXX ( worst, [@pp] )
VAR wAfter  = base + wPP
VAR worstNR = TOPN ( 1, real, [@nr], ASC )
VAR wnName  = MAXX ( worstNR, FactData[Sens Shock] )
VAR wnV     = MAXX ( worstNR, [@nr] )
VAR nBreach = COUNTROWS ( FILTER ( real, base + [@pp] < hardLim ) )
VAR allTxt  = CONCATENATEX ( real, "- " & FactData[Sens Shock] & ": " & FORMAT ( [@pp], fmtP ) & " pp", nl, [@pp], ASC )
VAR pageText =
    "SENSITIVITIES " & y & ": " & UPPER ( who ) & nl & nl
    & IF ( COUNTROWS ( real ) = 0, "No sensitivity results reported.",
           "Solvency ratio before shocks: " & FORMAT ( base, "0.0" ) & "%." & nl
           & "Most severe shock: " & wName & " (" & FORMAT ( wPP, fmtP ) & " pp), ratio after shock "
           & FORMAT ( wAfter, "0.0" ) & "%"
           & IF ( wAfter < hardLim, ": BELOW the " & hardLim & "% threshold.",
                  IF ( wAfter < softLim, ": between the " & hardLim & "% and " & softLim & "% thresholds.",
                       ": still above the " & softLim & "% threshold." ) ) & nl
           & IF ( nBreach > 0, nBreach & " shock(s) would bring the ratio below " & hardLim & "%.",
                  "No shock brings the ratio below " & hardLim & "%." ) & nl
           & IF ( wnV < 0, "Largest hit on net result: " & wnName & " (" & FORMAT ( wnV, fmt ) & ").", "" )
           & nl & nl & "Impact on the solvency ratio:" & nl & allTxt )
VAR sh  = SELECTEDVALUE ( FactData[Sens Shock] )
VAR pp  = [Sens Ratio pp]
VAR aft = base + pp
VAR rowText =
    UPPER ( sh ) & " " & y & ": " & who & nl & nl
    & "Impact on net result: " & FORMAT ( [Sens NR] + 0, fmt ) & nl
    & "Impact on own funds: " & FORMAT ( [Sens EOF] + 0, fmt ) & nl
    & "Impact on solvency ratio: " & FORMAT ( pp, fmtP ) & " pp, from " & FORMAT ( base, "0.0" )
    & "% to " & FORMAT ( aft, "0.0" ) & "%"
    & IF ( aft < hardLim, ": BELOW the " & hardLim & "% threshold.", "." )
RETURN IF ( HASONEVALUE ( FactData[Sens Shock] ), rowText, pageText )


=====================================================================
 3. FREE TANGIBLE CAPITAL (FTC)
=====================================================================

>> Что показываем: FTC по компании и по годам, сравнение с прошлым
>> планом и сравнение компаний между собой.
>> FTC — это остаток (stock), поэтому в комментарии берётся
>> последний год, а не сумма лет (useLastYear = TRUE).

--- 3.1 Визуалы ---
  a) Срез Entity (синхронизирован).
  b) Матрица Current: Righe = Item Label, Colonne = Year, Valori = Value Current.
     Totali di riga = Disattivato.
  c) Матрица Delta: то же, Valori = Delta CvP (условный цвет: <0 красный).
  d) Линейный график "FTC trend":
       Asse X = Year, Asse Y = Value Current и Value Previous
       Фильтр визуала: Item Label "contiene" Free Tangible Capital
  e) Столбчатый график "FTC by company" (последний год):
       Asse X = FactData[Entity], Asse Y = Value Current
       Фильтр визуала: Item Label "contiene" Free Tangible Capital, Year = последний
       Выделить срез Entity > Modifica interazioni > над этим графиком "Nessuno",
       чтобы график всегда показывал все компании.
  f) Комментарий: FTC Comment (ниже).

--- 3.2 Комментарий ---
>> Это универсальная мера. Её же используем для HMS (раздел 4),
>> меняются только две первые строки: tp и useLastYear.

FTC Comment =
VAR tp = "Free Tangible Capital"
VAR useLastYear = TRUE ()
VAR fmt  = "#,##0"
VAR fmtD = "+#,##0;-#,##0;0"
VAR nl   = UNICHAR ( 10 )
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ),
              FactData[Topic] = tp, NOT ISBLANK ( FactData[Value] ) )
VAR y0 = MINX ( yrs, FactData[Year] )
VAR y1 = MAXX ( yrs, FactData[Year] )
VAR period = IF ( y0 = y1, "in " & y1, y0 & "-" & y1 )
VAR scopeTxt = IF ( useLastYear, "in " & y1, period )
VAR yrSet = FILTER ( ALL ( FactData[Year] ),
                IF ( useLastYear, FactData[Year] = y1, FactData[Year] >= y0 && FactData[Year] <= y1 ) )
VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ), FactData[Topic] = tp )
VAR ent  = CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ), FactData[Topic] = tp )
VAR who  = IF ( nEnt = 1, IF ( ISBLANK ( ent ) || ent = "", "", ent ), nEnt & " companies" )
VAR itemTbl =
    CALCULATETABLE (
        ADDCOLUMNS ( SUMMARIZE ( FactData, DimItem[Item Label] ),
            "@c", [Value Current],
            "@p", [Value Previous],
            "@d", SUMX ( VALUES ( FactData[Year] ), [Delta CvP] ) ),
        REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ),
        FactData[Topic] = tp, FactData[IsTotal] = FALSE (), yrSet )
VAR totRow = CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[Item] ),
                 FactData[Topic] = tp, FactData[IsTotal] = TRUE (), yrSet )
VAR active  = FILTER ( itemTbl, [@c] <> 0 )
VAR top3Txt = CONCATENATEX ( TOPN ( 3, active, ABS ( [@c] ), DESC ),
                  DimItem[Item Label] & " (" & FORMAT ( [@c], fmt ) & ")", "; ", ABS ( [@c] ), DESC )
VAR hasPrev = COUNTROWS ( FILTER ( itemTbl, [@p] <> 0 ) ) > 0
VAR chg     = FILTER ( itemTbl, [@d] <> 0 )
VAR chgTxt  = CONCATENATEX ( TOPN ( 3, chg, ABS ( [@d] ), DESC ),
                  DimItem[Item Label] & " (" & FORMAT ( [@d], fmtD ) & ")", "; ", ABS ( [@d] ), DESC )
VAR pageText =
    UPPER ( tp ) & IF ( who <> "", ": " & UPPER ( who ), "" ) & ", " & scopeTxt & nl & nl
    & IF ( COUNTROWS ( active ) = 0, "No amounts in the current plan.",
           IF ( NOT ISBLANK ( totRow ), "Total: " & FORMAT ( totRow, fmt ) & ".",
                "Sum of all lines: " & FORMAT ( SUMX ( itemTbl, [@c] ), fmt ) & " (" & COUNTROWS ( active ) & " lines with amounts)." )
           & nl & "Largest lines: " & top3Txt & "." )
    & nl & nl
    & IF ( hasPrev,
           IF ( COUNTROWS ( chg ) = 0, "Fully in line with the previous plan.",
                "Main changes vs previous plan: " & chgTxt & "." ),
           "No previous plan available for comparison." )
VAR lbl = SELECTEDVALUE ( DimItem[Item Label] )
VAR rowYrs = ADDCOLUMNS ( yrs, "@c", CALCULATE ( [Value Current], FactData[Topic] = tp ),
                               "@p", CALCULATE ( [Value Previous], FactData[Topic] = tp ) )
VAR rv0 = MAXX ( FILTER ( rowYrs, FactData[Year] = y0 ), [@c] )
VAR rv1 = MAXX ( FILTER ( rowYrs, FactData[Year] = y1 ), [@c] )
VAR rCommon = FILTER ( rowYrs, [@p] <> 0 )
VAR rD  = SUMX ( rCommon, [@c] - [@p] )
VAR byYear = CONCATENATEX ( rowYrs, FactData[Year] & ": " & FORMAT ( [@c] + 0, fmt ), " | ", FactData[Year], ASC )
VAR rowText =
    UPPER ( lbl ) & IF ( who <> "", ": " & who, "" ) & nl & nl
    & IF ( y0 = y1,
           "Value in " & y1 & ": " & FORMAT ( rv1 + 0, fmt ) & ".",
           "By year: " & byYear & "." & nl
           & IF ( useLastYear,
                  "Change " & y0 & "-" & y1 & ": " & FORMAT ( rv1 - rv0, fmtD )
                  & IF ( rv0 <> 0, " (" & FORMAT ( DIVIDE ( rv1 - rv0, ABS ( rv0 ) ), "+0%;-0%;0%" ) & ")", "" ) & ".",
                  "Total " & period & ": " & FORMAT ( SUMX ( rowYrs, [@c] ) + 0, fmt ) & "." ) )
    & IF ( COUNTROWS ( rCommon ) > 0,
           nl & "Versus previous plan: " & FORMAT ( rD, fmtD )
           & IF ( rD > 0, " (higher)", IF ( rD < 0, " (lower)", " (unchanged)" ) ) & ".",
           nl & "No previous plan for this line." )
RETURN IF ( HASONEVALUE ( DimItem[Item Label] ), rowText, pageText )


=====================================================================
 4. HMS IMPACTS
=====================================================================

>> Что показываем: влияние мер HMS по строкам и годам, текущий план
>> против прошлого. Структура как у Capital Needs, поэтому проще
>> всего скопировать страницу Capital Needs (Duplica pagina) и
>> поменять фильтр страницы на Topic = HMS Impacts.

--- 4.1 Визуалы ---
  a) Три матрицы (как Capital Needs): Value Current / Value Previous / Delta CvP.
     Righe = Item Label, Colonne = Year, Totali di riga = Disattivato.
     >> Если в шаге 0 увидишь компании в Entity — добавь срез Entity.
  b) Istogramma raggruppato "Current vs Previous by year":
       Asse X = Year, Asse Y = Value Current и Value Previous
       Фильтр визуала: IsTotal = Vero (если в листе есть строка Total),
       иначе без фильтра (сумма всех строк).
  c) Водопад "What drives the total" (за один год):
       Categoria = Item Label, Asse Y = Value Current
       Фильтр визуала: IsTotal = Falso
       Рядом срез Year (Selezione singola), через Modifica interazioni
       влияет только на водопад.
  d) Комментарий: HMS Comment.

--- 4.2 Комментарий ---
>> Скопируй ВСЮ меру FTC Comment, переименуй в HMS Comment
>> и поменяй ТОЛЬКО первые две строки после знака "=":

HMS Comment =
VAR tp = "HMS Impacts"
VAR useLastYear = FALSE ()
   ... (дальше всё как в FTC Comment, без изменений) ...

>> FALSE = это потоки, комментарий считает сумму за период.


=====================================================================
 5. ЧЕК-ЛИСТ ПОСЛЕ КАЖДОЙ СТРАНИЦЫ
=====================================================================
  [ ] Фильтр страницы Topic стоит, лишних фильтров Item нет
  [ ] 2-3 цифры сверены с Excel (Scale = Millions)
  [ ] Проверки (Check) = 0
  [ ] Комментарий: ничего не выбрано -> сводка; клик по строке -> строка
  [ ] Срезы на тех же местах, что и на других страницах
  [ ] Ctrl+S и копия файла с датой

ЧТО ПРИСЛАТЬ МНЕ, ЕСЛИ НЕ РАБОТАЕТ
  - Фото Dati > FactData с фильтром Topic (колонки Block, Section, Item, Measure)
  - Фото ошибки DAX целиком (строка + подсказка)
  - Фото страницы с выбранными срезами
