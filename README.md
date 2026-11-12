=====================================================================
 CDD -> POWER BI  |  ЧАСТЬ A: РЫНОЧ
                  |  ЧАСТЬ B: ДИЗАЙН CDD КАК В CPR + ГРАФИКИ
                  |  ЧАСТЬ C: ОПТИМИЗАЦИЯ (меньше мер и колонок, результат тот же)
 Строки с ">>" не копировать. Копировать только код.
 Работать в КОПИИ .pbix (например CDD_v4_market.pbix).
=====================================================================

>> ЧЕСТНО О ДАННЫХ (важно для презентации)
>> - ECB Data Portal: официально, бесплатно, без ключа, есть API -> подключаем сразу.
>>   Даёт: кривую доходности гособлигаций еврозоны (AAA), доходности 10Y по странам
>>   (BTP, Bund, Bonos ...), €STR, Euribor 3M, ставку депозитов ECB, инфляцию HICP по странам,
>>   EUR/USD, индекс Euro Stoxx 50.
>> - "Curva Euro Swap" (EUR IRS) ECB бесплатно НЕ публикует. Варианты:
>>     сейчас  -> кривая AAA ECB как прокси (подписываем честно: "EUR govt AAA curve, ECB");
>>     лучше   -> кривая EIOPA RFR (официальная, строится из свопов, это и есть кривая Solvency II,
>>                ежемесячно, Excel-файл с сайта EIOPA) — ШАГ A9;
>>     потом   -> Bloomberg (нужна лицензия / Data License), если тема пойдёт дальше.
>> - Yahoo Finance: неофициальный, без гарантий и с правовыми оговорками -> для Generali не советую.
>> - Частота: доходности 10Y по странам — МЕСЯЧНЫЕ (среднее за месяц), кривая AAA и €STR — дневные.
>>   Для обзора капитала этого достаточно; "real-time" будет при Bloomberg.


#####################################################################
#####################################################################
 ЧАСТЬ A. РЫНОЧНЫЕ ДАННЫЕ ECB
#####################################################################
#####################################################################

#####################################################################
 A1. ПРОВЕРКА ДОСТУПА (2 минуты, ДО Power BI)
#####################################################################
>> Открой в браузере на рабочем ПК (должна скачаться/открыться таблица CSV):
>>   https://data-api.ecb.europa.eu/service/data/IRS/M.IT.L.L40.CI.0000.EUR.N.Z?format=csvdata&startPeriod=2025-01
>> Открылось -> всё ок.
>> Ошибка/блокировка -> прокси Generali закрывает сайт: нужен запрос в IT
>> ("whitelist data-api.ecb.europa.eu for Power BI"). Без этого Часть A не заработает.


#####################################################################
 A2. ПАРАМЕТР MktStart
#####################################################################
>> Trasforma dati > Gestisci parametri > Nuovo:
>>   Nome: MktStart   Tipo: Testo   Valore: 2021-01
>> (с какого месяца грузить историю)


#####################################################################
 A3. ФУНКЦИЯ fnECB  (Query vuota > Editor avanzato > вставить > имя fnECB)
#####################################################################

(flow as text, key as text, series as text) as table =>
let
    Src   = Csv.Document(
                Web.Contents("https://data-api.ecb.europa.eu",
                    [RelativePath = "service/data/" & flow & "/" & key,
                     Query = [format = "csvdata", startPeriod = MktStart]]),
                [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Hdr   = Table.PromoteHeaders(Src, [PromoteAllScalars = true]),
    Keep  = Table.SelectColumns(Hdr, {"KEY", "REF_AREA", "TIME_PERIOD", "OBS_VALUE"}, MissingField.UseNull),
    NoNa  = Table.SelectRows(Keep, each [OBS_VALUE] <> null and [OBS_VALUE] <> "" and [OBS_VALUE] <> "NaN"),
    Out   = Table.FromRecords(List.Transform(Table.ToRecords(NoNa), (r) => [
                Series  = series,
                Country = if r[REF_AREA] = null or r[REF_AREA] = "U2" then "EA" else r[REF_AREA],
                Tenor   = if flow = "YC" then Text.AfterDelimiter(r[KEY], "SR_") else null,
                Date    = Date.From(if Text.Length(r[TIME_PERIOD]) = 7 then r[TIME_PERIOD] & "-01" else r[TIME_PERIOD]),
                Value   = Number.From(r[OBS_VALUE], "en-US")
            ]),
            type table [Series = text, Country = text, Tenor = text, Date = date, Value = number])
in
    Out


#####################################################################
 A4. ЗАПРОС MarketData  (Query vuota > вставить > имя MarketData, load ON)
#####################################################################
>> Одна длинная таблица: Series | Country | Tenor | Date | Value (в %).
>> Страны: IT DE ES FR AT PT GR NL BE — допиши нужные через "+".

let
    Ctry = "IT+DE+ES+FR+AT+PT+GR+NL+BE",
    L = {
        fnECB("YC",  "B.U2.EUR.4F.G_N_A.SV_C_YM.SR_1Y+SR_2Y+SR_3Y+SR_5Y+SR_7Y+SR_10Y+SR_15Y+SR_20Y+SR_30Y", "EUR_AAA_CURVE"),
        fnECB("IRS", "M." & Ctry & ".L.L40.CI.0000.EUR.N.Z",          "GOV10Y"),
        fnECB("EST", "B.EU000A2X2A25.WT",                             "ESTR"),
        fnECB("FM",  "M.U2.EUR.RT.MM.EURIBOR3MD_.HSTA",               "EURIBOR3M"),
        fnECB("FM",  "D.U2.EUR.4F.KR.DFR.LEV",                        "ECB_DFR"),
        fnECB("ICP", "M." & Ctry & ".N.000000.4.ANR",                 "HICP"),
        fnECB("EXR", "D.USD.EUR.SP00.A",                              "EURUSD"),
        try fnECB("FM", "M.U2.EUR.DS.EI.DJES50I.HSTA", "EUROSTOXX50") otherwise #table(
            type table [Series = text, Country = text, Tenor = text, Date = date, Value = number], {})
    },
    All   = Table.Combine(L),
    Month = Table.AddColumn(All, "Month", each Date.StartOfMonth([Date]), type date),
    TOrd  = Table.AddColumn(Month, "TenorYears", each
                if [Tenor] = null then null else Number.From(Text.Remove([Tenor], {"Y"})), Int64.Type)
in
    TOrd

>> Если одна из строк даёт ошибку (сайт поменял ключ) — закомментируй её "//" и напиши мне.
>> Privacy: если спросит уровень — Pubblico для ECB (или оставь "Ignora i livelli di privacy").
>> Проверка: Series = 8 значений, у GOV10Y и HICP есть Country IT/DE/ES..., у кривой Tenor 1..30.


#####################################################################
 A5. ТАБЛИЦЫ ДЛЯ ВЫБОРА  (Modellazione > Nuova tabella)
#####################################################################

DimMktCountry =
DATATABLE ( "Code", STRING, "Country", STRING, "Order", INTEGER,
    { { "IT", "Italy", 1 }, { "ES", "Spain", 2 }, { "FR", "France", 3 }, { "AT", "Austria", 4 },
      { "PT", "Portugal", 5 }, { "GR", "Greece", 6 }, { "NL", "Netherlands", 7 }, { "BE", "Belgium", 8 },
      { "DE", "Germany", 9 } } )

>> СВЯЗИ НЕ СОЗДАВАТЬ (ни с MarketData, ни с FactData): страна передаётся через меру.
>> MarketData[TenorYears], [Value] > Riepilogo: Non riepilogare.
>> MarketData[Tenor] > Ordina per colonna > TenorYears.

>> Параметры сценария (Modellazione > Nuovo parametro > Intervallo numerico):
>>   Nome "Shock Spread"   Min -100  Max 300  Incremento 10  Predefinito 0   (bp)
>>   Nome "Shock Rates"    Min -100  Max 100  Incremento 10  Predefinito 0   (bp)
>>   Nome "Shock Equity"   Min -40   Max 20   Incremento 5   Predefinito 0   (%)
>>   "Aggiungi filtro dei dati a questa pagina" = Sì
>> Power BI создаст меры [Shock Spread Valore], [Shock Rates Valore], [Shock Equity Valore].


#####################################################################
 A6. МЕРЫ РЫНКА  (Nuova misura, папка "Market")
#####################################################################

Mkt Country = SELECTEDVALUE ( DimMktCountry[Code], "IT" )

Mkt Country Name = SELECTEDVALUE ( DimMktCountry[Country], "Italy" )

>> Дата сравнения = начало плана (конец года Actual). Поменяй при новом цикле.
Mkt Ref Date = DATE ( 2025, 12, 31 )

>> ---------- последние значения (для карточек; работают в любом визуале) ----------
Gov10Y Last =
VAR c = [Mkt Country]
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "GOV10Y", MarketData[Country] = c )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

Bund10Y Last =
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "GOV10Y", MarketData[Country] = "DE" )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

AAA10Y Last =
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "EUR_AAA_CURVE", MarketData[Tenor] = "10Y" )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

ESTR Last =
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "ESTR" )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

>> Ещё 4 меры по ТОМУ ЖЕ шаблону, что ESTR Last — меняешь только Series:
>>   Euribor3M Last  -> "EURIBOR3M"
>>   ECB DFR Last    -> "ECB_DFR"
>>   EURUSD Last     -> "EURUSD"
>>   EuroStoxx Last  -> "EUROSTOXX50"

HICP Last =
VAR c = [Mkt Country]
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "HICP", MarketData[Country] = c )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

Mkt As Of =
"ECB Data Portal · last data: " & FORMAT (
    CALCULATE ( MAX ( MarketData[Date] ), REMOVEFILTERS ( MarketData ), MarketData[Series] = "GOV10Y" ), "mmm yyyy" )

>> ---------- спреды (в bp) ----------
Spread vs Bund bp = ( [Gov10Y Last] - [Bund10Y Last] ) * 100

Spread vs AAA bp = ( [Gov10Y Last] - [AAA10Y Last] ) * 100

Spread vs Bund Ref bp =
VAR c  = [Mkt Country]
VAR m  = DATE ( YEAR ( [Mkt Ref Date] ), MONTH ( [Mkt Ref Date] ), 1 )
VAR g  = CALCULATE ( MAX ( MarketData[Value] ), REMOVEFILTERS ( MarketData ),
             MarketData[Series] = "GOV10Y", MarketData[Country] = c, MarketData[Month] = m )
VAR b  = CALCULATE ( MAX ( MarketData[Value] ), REMOVEFILTERS ( MarketData ),
             MarketData[Series] = "GOV10Y", MarketData[Country] = "DE", MarketData[Month] = m )
RETURN IF ( NOT ISBLANK ( g ) && NOT ISBLANK ( b ), ( g - b ) * 100 )

Spread Move bp = [Spread vs Bund bp] - [Spread vs Bund Ref bp]

AAA10Y Move bp =
VAR m  = DATE ( YEAR ( [Mkt Ref Date] ), MONTH ( [Mkt Ref Date] ), 1 )
VAR r  = CALCULATE ( AVERAGE ( MarketData[Value] ), REMOVEFILTERS ( MarketData ),
             MarketData[Series] = "EUR_AAA_CURVE", MarketData[Tenor] = "10Y", MarketData[Month] = m )
RETURN IF ( NOT ISBLANK ( r ), ( [AAA10Y Last] - r ) * 100 )

EuroStoxx Move % =
VAR m  = DATE ( YEAR ( [Mkt Ref Date] ), MONTH ( [Mkt Ref Date] ), 1 )
VAR r  = CALCULATE ( MAX ( MarketData[Value] ), REMOVEFILTERS ( MarketData ),
             MarketData[Series] = "EUROSTOXX50", MarketData[Month] = m )
RETURN IF ( NOT ISBLANK ( r ), DIVIDE ( [EuroStoxx Last], r ) - 1 )

>> ---------- для графиков (работают по оси Month / Tenor, БЕЗ REMOVEFILTERS) ----------
Gov10Y M =
CALCULATE ( AVERAGE ( MarketData[Value] ), MarketData[Series] = "GOV10Y", MarketData[Country] = [Mkt Country] )

Bund10Y M =
CALCULATE ( AVERAGE ( MarketData[Value] ), MarketData[Series] = "GOV10Y", MarketData[Country] = "DE" )

AAA10Y M =
CALCULATE ( AVERAGE ( MarketData[Value] ), MarketData[Series] = "EUR_AAA_CURVE", MarketData[Tenor] = "10Y" )

Spread Bund M bp = IF ( NOT ISBLANK ( [Gov10Y M] ), ( [Gov10Y M] - [Bund10Y M] ) * 100 )

Spread AAA M bp = IF ( NOT ISBLANK ( [Gov10Y M] ), ( [Gov10Y M] - [AAA10Y M] ) * 100 )

HICP M =
CALCULATE ( AVERAGE ( MarketData[Value] ), MarketData[Series] = "HICP", MarketData[Country] = [Mkt Country] )

>> Кривая на последнюю дату и на дату сравнения (ось = MarketData[Tenor])
Curve Last =
VAR d = CALCULATE ( MAX ( MarketData[Date] ), REMOVEFILTERS ( MarketData ), MarketData[Series] = "EUR_AAA_CURVE" )
RETURN CALCULATE ( AVERAGE ( MarketData[Value] ), MarketData[Series] = "EUR_AAA_CURVE", MarketData[Date] = d )

Curve Ref =
VAR d = CALCULATE ( MAX ( MarketData[Date] ), REMOVEFILTERS ( MarketData ),
            MarketData[Series] = "EUR_AAA_CURVE", MarketData[Date] <= [Mkt Ref Date] )
RETURN CALCULATE ( AVERAGE ( MarketData[Value] ), MarketData[Series] = "EUR_AAA_CURVE", MarketData[Date] = d )

>> ФОРМАТЫ: *Last, *M, Curve -> "0.00" и подпись " %" в заголовке; *bp -> "+0;-0;0";
>> EuroStoxx Move % -> "+0.0%;-0.0%".


#####################################################################
 A7. СЦЕНАРИЙ: связь рынка с чувствительностями CDD  (папка "Scenario")
#####################################################################
>> Идея: в CDD каждая компания сообщает влияние шоков на Solvency ratio (pp):
>> "+50 bps all govies spread", "+50 bps Interest Rates", "-50 bps Interest Rates", "-25% Equity".
>> Масштабируем их линейно на текущее движение рынка или на шок из ползунка:
>>     влияние = чувствительность × (шок / размер шока в CDD)
>> Это ОЦЕНКА первого порядка (без выпуклости, VA, взаимодействия рисков) — так и подписываем.
>> Используются меры страницы Sensitivities: [Sens Ratio pp], [Sens Base Ratio] и колонка FactData[Sens Shock].
>> Если какой-то шок в файле назван иначе — поменяй слово в CONTAINSSTRING.

Sens Year =
CALCULATE ( MIN ( FactData[Year] ), REMOVEFILTERS ( DimPeriod ),
    FactData[Topic] = "Sensitivities", NOT ISBLANK ( FactData[Sens Shock] ) )

Sens Govies pp =
VAR y = [Sens Year]
RETURN CALCULATE ( [Sens Ratio pp], FactData[Year] = y,
    KEEPFILTERS ( FILTER ( ALL ( FactData[Sens Shock] ), CONTAINSSTRING ( FactData[Sens Shock], "govies" ) ) ) )

Sens IR Up pp =
VAR y = [Sens Year]
RETURN CALCULATE ( [Sens Ratio pp], FactData[Year] = y,
    KEEPFILTERS ( FILTER ( ALL ( FactData[Sens Shock] ),
        CONTAINSSTRING ( FactData[Sens Shock], "+50" ) && CONTAINSSTRING ( FactData[Sens Shock], "Interest" ) ) ) )

Sens IR Down pp =
VAR y = [Sens Year]
RETURN CALCULATE ( [Sens Ratio pp], FactData[Year] = y,
    KEEPFILTERS ( FILTER ( ALL ( FactData[Sens Shock] ),
        CONTAINSSTRING ( FactData[Sens Shock], "-50" ) && CONTAINSSTRING ( FactData[Sens Shock], "Interest" ) ) ) )

Sens Equity pp =
VAR y = [Sens Year]
RETURN CALCULATE ( [Sens Ratio pp], FactData[Year] = y,
    KEEPFILTERS ( FILTER ( ALL ( FactData[Sens Shock] ), CONTAINSSTRING ( FactData[Sens Shock], "Equity" ) ) ) )

Scn Base SR =
VAR y = [Sens Year]
RETURN CALCULATE ( [Sens Base Ratio], FactData[Year] = y ) / 100

>> ---------- шоки: из ползунков (Manual) или из рынка (Market) ----------
Scn Spread bp  = [Shock Spread Valore]
Scn Rates bp   = [Shock Rates Valore]
Scn Equity pct = [Shock Equity Valore]

Impact Spread pp = DIVIDE ( [Scn Spread bp], 50 ) * [Sens Govies pp]

Impact Rates pp =
VAR s  = [Scn Rates bp]
VAR up = [Sens IR Up pp]
VAR dn = COALESCE ( [Sens IR Down pp], -up )
RETURN IF ( s >= 0, s / 50 * up, -s / 50 * dn )

Impact Equity pp = DIVIDE ( [Scn Equity pct], -25 ) * [Sens Equity pp]

Impact Total pp = [Impact Spread pp] + [Impact Rates pp] + [Impact Equity pp]

Scn SR = IF ( NOT ISBLANK ( [Scn Base SR] ), [Scn Base SR] + [Impact Total pp] / 100 )

>> Тот же расчёт на ФАКТИЧЕСКОМ движении рынка с даты сравнения
Mkt Impact pp =
VAR sp = [Spread Move bp]
VAR rt = [AAA10Y Move bp]
VAR eq = [EuroStoxx Move %] * 100
VAR up = [Sens IR Up pp]
VAR dn = COALESCE ( [Sens IR Down pp], -up )
RETURN
    DIVIDE ( sp, 50 ) * [Sens Govies pp]
    + IF ( rt >= 0, rt / 50 * up, -rt / 50 * dn )
    + DIVIDE ( eq, -25 ) * [Sens Equity pp]

Mkt SR = IF ( NOT ISBLANK ( [Scn Base SR] ), [Scn Base SR] + [Mkt Impact pp] / 100 )

Scn Status =
VAR r   = [Scn SR]
VAR red = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "SOL_RED" )
VAR amb = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "SOL_AMBER" )
RETURN SWITCH ( TRUE (), ISBLANK ( r ), BLANK (), r < red, 3, r < amb, 2, 1 )

Scn Color = SWITCH ( [Scn Status], 1, "#2E7D32", 2, "#E0A100", 3, "#A50E0E", "#9E9E9E" )

Scn Comment =
VAR nl = UNICHAR ( 10 )
RETURN
"MARKET SINCE " & UPPER ( FORMAT ( [Mkt Ref Date], "mmm yyyy" ) ) & " – " & UPPER ( [Mkt Country Name] ) & nl
& "10Y spread vs Bund: " & FORMAT ( [Spread vs Bund bp], "0" ) & " bp (" & FORMAT ( [Spread Move bp], "+0;-0;0" ) & " bp)." & nl
& "EUR 10Y AAA rate: " & FORMAT ( [AAA10Y Last], "0.00" ) & "% (" & FORMAT ( [AAA10Y Move bp], "+0;-0;0" ) & " bp)." & nl
& IF ( NOT ISBLANK ( [EuroStoxx Move %] ), "Euro Stoxx 50: " & FORMAT ( [EuroStoxx Move %], "+0.0%;-0.0%" ) & "." & nl, "" )
& IF ( HASONEVALUE ( FactData[Entity] ) && NOT ISBLANK ( [Mkt SR] ),
     nl & SELECTEDVALUE ( FactData[Entity] ) & ": estimated solvency ratio " & FORMAT ( [Scn Base SR], "0%" )
     & " → " & FORMAT ( [Mkt SR], "0%" ) & " (" & FORMAT ( [Mkt Impact pp], "+0.0;-0.0" ) & " pp) at current market levels.",
     nl & "Select a company to see the estimated impact on its solvency ratio." )
& nl & nl & "First-order estimate from reported sensitivities; not a full recalculation."

>> Почему по компании: коэффициенты разных компаний складывать нельзя (как в CPR).
>> В таблице по компаниям (A8, визуал 5) каждая строка считается отдельно.


#####################################################################
 A8. СТРАНИЦА "Market & Scenario"  (1280 x 720, дизайн как в CPR)
#####################################################################
>> Шапка #7A0F0C: "Market indicators & scenario" + Scheda с [Mkt As Of].
>> Слайсер DimMktCountry[Country] (Elenco a discesa, Selezione singola), сорт. по Order.
>>
>> 1) KPI-ряд (6 карточек 190x80): Gov10Y Last ("10Y govt yield, %"), Spread vs Bund bp,
>>    Spread Move bp ("Δ since ref."), ESTR Last, ECB DFR Last, HICP Last.
>>    Цвет Spread Move: Regole >0 #A50E0E, <0 #2E7D32.
>> 2) "EUR yield curve": Grafico a linee; Asse X = MarketData[Tenor];
>>    Asse Y = Curve Last, Curve Ref. Цвета: Last #C21B17, Ref #9E9E9E (tratteggio).
>>    Подпись: "EUR govt AAA curve (ECB) – latest vs reference date".
>> 3) "Spread evolution": Grafico a linee; Asse X = MarketData[Month] (Gerarchia off);
>>    Asse Y = Spread Bund M bp, Spread AAA M bp.
>>    Analisi > Linea X costante = Mkt Ref Date (fx) "Plan reference".
>> 4) "Rates": Grafico a linee, X = Month, Y = Gov10Y M, Bund10Y M, AAA10Y M.
>> 5) "Scenario by company": Matrice
>>      Righe = FactData[Entity]
>>      Valori = Scn Base SR | Impact Spread pp | Impact Rates pp | Impact Equity pp | Scn SR | Mkt SR
>>      Scn SR: Colore carattere > Valore campo > Scn Color.
>>      Фильтр визуала: Scn Base SR "non è vuoto".
>>    Над ней три ползунка: Shock Spread, Shock Rates, Shock Equity.
>> 6) Key messages справа: Scn Comment.
>> 7) Сноска: "Source: ECB Data Portal (monthly 10Y yields = monthly averages). Scenario = linear
>>    scaling of reported CDD sensitivities. Indicative only."
>>
>> ВСТРОИТЬ В СУЩЕСТВУЮЩИЕ СТРАНИЦЫ (просто, но эффектно):
>>   - Solvency / Sensitivities: карточка Mkt SR + Spread Move bp рядом с матрицей.
>>   - Summary Total: строка в Executive Summary (см. ниже) и карточка "Spread vs Bund".
>>   - SAA: карточки Gov10Y Last и AAA10Y Last (контекст доходности портфеля).
>>
>> Строка рынка в Executive Summary (добавь в конце RETURN меры Executive Summary):
>>   & nl & nl & "MARKET" & nl & b & [Mkt Country Name] & " 10Y spread vs Bund "
>>   & FORMAT ( [Spread vs Bund bp], "0" ) & " bp (" & FORMAT ( [Spread Move bp], "+0;-0;0" )
>>   & " bp since " & FORMAT ( [Mkt Ref Date], "mmm yyyy" ) & ")."


#####################################################################
 A9. (СЛЕДУЮЩИЙ ШАГ) КРИВАЯ EIOPA RFR = "CURVA EURO SWAP" для Solvency II
#####################################################################
>> 1) eiopa.europa.eu > Risk-free interest rate term structures > последний месяц > скачай zip,
>>    распакуй "EIOPA_RFR_YYYYMMDD_Term_Structures.xlsx" в папку MarketFiles.
>> 2) Пришли мне фото листа "RFR_spot_no_VA" (видно строки сроков и колонку Euro) —
>>    напишу парсер (как fnReadSheet), чтобы кривая и "Spread Italy vs EUR swap" считались от неё.
>> Это будет официальная кривая из свопов — то, что было в CDD как "Curva Euro Swap".


#####################################################################
#####################################################################
 ЧАСТЬ B. ДИЗАЙН CDD КАК В CPR + ГРАФИКИ
#####################################################################
#####################################################################

>> B1. ДВИЖОК СТРОК (как в CPR): добавь в FactData CDD хвост LineKey / Line / LineOrder /
>>     BlockOrder из CPR (тот же код: WithKey + LK..Final; в CDD последний шаг Final3 —
>>     подставь его вместо WithSrc). И колонку PeriodSort из CPR, но с порядком CDD:

PeriodSort =
FactData[Year] * 10
    + SWITCH ( FactData[PeriodType], "Actual", 1, "Half-Year", 2, "Pro Forma SII", 3,
               "Forecast", 4, "Year", 5, "Plan", 6, 9 )

>>     Ordina per colonna: Period -> PeriodSort, Line -> LineOrder, Block -> BlockOrder.
>>     Меры формата из CPR (Line Kind, CPR Cell, Cell BG, Cell Font) работают и в CDD:
>>     в Cell BG замени условие isCY на "первый год плана":
>>        VAR isCY = SELECTEDVALUE ( FactData[PeriodType] ) = "Plan"
>>                   && SELECTEDVALUE ( FactData[Year] ) = CALCULATE ( MIN ( FactData[Year] ),
>>                        REMOVEFILTERS ( FactData ), FactData[PeriodType] = "Plan" )
>>     В CDD нет RatioFlag — создай колонку RatioFlag (код из CPR, ШАГ 3).
>>
>> B2. ШАБЛОН СТРАНИЦЫ (одинаково на всех):
>>   ┌──────────────────────────────────────────────────────────────────────┐
>>   │ ШАПКА #7A0F0C: название листа + Page Subtitle | слайсеры Entity, Scale │
>>   ├──────────────────────────────┬─────────────────────────┬─────────────┤
>>   │ МАТРИЦА "как в Excel"          │ ГРАФИК (см. B3)          │ KEY         │
>>   │ X16 Y72 W560 H400              │ X584 Y72 W380 H400       │ MESSAGES    │
>>   ├──────────────────────────────┴─────────────────────────┤ X972 W292   │
>>   │ KPI-КАРТОЧКИ (3–4 шт., H 90) + 2-й график / таблица        │ [X Comment] │
>>   │ X16 Y480 W948 H224                                          │             │
>>   └────────────────────────────────────────────────────────┴─────────────┘
>>
>> B3. ГРАФИКИ ПО СТРАНИЦАМ (меры у тебя уже есть; если имя другое — бери меру из матрицы страницы)
>>  Страница                 График 1                                      График 2 / карточки
>>  -----------------------  --------------------------------------------  ------------------------------
>>  Solvency, Distr. & FTC   Линии: SOL Ratio + SOL Ratio Prev по годам;     Карточки: SOL Min Ratio,
>>                           Analisi > Linea Y costante fx = 'Thresholds'     FTC последний год, Mkt SR
>>                           SOL_RED и SOL_AMBER (красная/жёлтая пунктир)
>>  Solvency AoM             Waterfall: Categoria = Line, Y = Value Current   Карточки: SR BoP, SR EoP, Δ pp
>>                           (только блок EOF), второй — SCR
>>  Sensitivities            Tornado: Barre, Y = Sens Shock, X = Sens Ratio   Матрица сценария из A8 (5)
>>                           pp, сорт. по возрастанию
>>  Free Tangible Capital    Waterfall по Line последнего года                Карточка FTC first→last
>>  Capital Needs            Colonne in cluster: Year × Current vs Previous    Карточка Total CN
>>  Total Remittance / RBC   Colonne in pila: Year × Dividends / Interim;     Карточка Total, Δ vs Prev
>>                           линия Previous
>>  Remittance Holding       Area/линия Cash position по годам; Linea Y=0     Карточки Cash start/end
>>  HMS                      Barre: Item × impact                              —
>>  SAA                      Barre in cluster: Asset class × Target vs        Карточки Gov10Y, AAA10Y
>>                           Current weight; Errori (min/max corridor)
>>  Life KPIs                Linea: NBV margin (LK Margin) по годам           Colonne: NBV по годам
>>  Non-Life KPIs            Linea: CoR по годам + Linea Y 100%               Colonne: Earned premiums
>>  Debt Management          Colonne: Reimbursements vs Issuances по годам    Таблица List_Debt
>>  Local RAF                Barre: Code × RAF Min SR; Linee Hard / Soft      Таблица List_RAF + RAF Color
>>  ORMT / Model Changes     Barre: ORMT/Change × pp                          Карточка total pp
>>  Initiatives              Grafico Gantt — нет; Tabella с INI Overdue       Карточка # overdue
>>
>>  Цвета графиков: основная серия #C21B17, Previous/Ref #9E9E9E, вторая #5C5C5C,
>>  статусы только #2E7D32 / #E0A100 / #A50E0E. Подписи данных — только на итогах.
>>
>> B4. Сделай сначала ОДНУ страницу (Solvency), пришли скрин, потом копируй шаблон.


#####################################################################
#####################################################################
 ЧАСТЬ C. ОПТИМИЗАЦИЯ: МЕНЬШЕ МЕР И КОЛОНОК, РЕЗУЛЬТАТ ТОТ ЖЕ
#####################################################################
#####################################################################
>> Честно: точное число смогу сказать, когда увижу список мер. Ожидаю −30…50% мер и
>> большую часть вычисляемых колонок. Пришли скрин левой панели "Visualizzazione TMDL"
>> или список мер (Visualizzazione modello > Modello semantico > Misure).
>>
>> C1. ОДИН ДВИЖОК СТРОК ВМЕСТО ~15 КОЛОНОК
>>     Колонки Line / LineOrder / BlockOrder из Power Query (B1) заменяют:
>>     LK Line, LK Order, NL Line, NL Order, MC Line, MC Order, MC RelRow, SAA Line, SAA Order,
>>     Sens Line, Sens Order, Sens RelRow и UNICHAR(8203)-трюки.
>>     Порядок: переставь матрицу страницы на Line -> сверь цифры -> удали старую колонку.
>>     Оставь только смысловые: Sens Shock, Sens Metric, LK Table, NL Group, MC Table
>>     (их тоже можно перенести в Power Query одной таблицей-маппингом ItemMap — шаг 2-й волны).
>>
>> C2. ГРУППА ВЫЧИСЛЕНИЙ вместо пар Current / Previous / Delta
>>     Visualizzazione modello > Modello semantico > Gruppi di calcolo > Nuovo gruppo di calcolo
>>     Имя "View", колонка "View". Элементы (Elemento di calcolo):

Current  = CALCULATE ( SELECTEDMEASURE (), FactData[Scenario] = "Current" )
Previous = CALCULATE ( SELECTEDMEASURE (), FactData[Scenario] = "Previous" )
Delta    = CALCULATE ( SELECTEDMEASURE (), FactData[Scenario] = "Current" )
           - CALCULATE ( SELECTEDMEASURE (), FactData[Scenario] = "Previous" )
Delta %  = VAR c = CALCULATE ( SELECTEDMEASURE (), FactData[Scenario] = "Current" )
           VAR p = CALCULATE ( SELECTEDMEASURE (), FactData[Scenario] = "Previous" )
           RETURN DIVIDE ( c - p, ABS ( p ) )

>>     Для Delta % в элементе: Espressione stringa di formato dinamica = "+0.0%;-0.0%".
>>     Использование: в матрице Colonne = View[View], Valori = базовая мера (например RBC Total).
>>     Тогда не нужны RBC Previous, SOL Ratio Prev, … Delta-меры (по одной на страницу).
>>     ВНИМАНИЕ: базовая мера не должна сама фильтровать Scenario (Value Current фильтрует ->
>>     используй [Value]). Power BI попросит "scoraggiare misure implicite" — согласись.
>>
>> C3. ОБЩИЕ HEADLINE / STATUS ЧЕРЕЗ SummaryIndex (как в CPR)
>>     Для простых страниц (TR, RBC, CN, RH, LK, NL, Debt, HMS) заменяют ~16 мер Headline/Status
>>     одной парой мер по KeyItem/Rule (CPR Headline / CPR Status, но Current vs Previous вместо
>>     Actual vs Forecast). Специальные оставить: SOL, AoM, FTC (проверки мостов), Sens, SAA, RAF.
>>     Это отдельный шаг — сделаю код, когда пришлёшь список мер.
>>
>> C4. УДАЛИТЬ / СКРЫТЬ
>>     DebugSAA, DimItem (после перехода на Line), старая Summary Overview, неиспользуемые
>>     *Comment HTML, CMP_* запросы (если IncludeCMP = FALSE и не планируешь) — Abilita caricamento Off.
>>     Перед удалением меры: правый клик > "Visualizza dipendenze" (если есть) или поиск по TMDL.
>>
>> C5. ИНСТРУМЕНТ: Visualizzazione TMDL (иконка TMDL слева у тебя уже есть)
>>     Позволяет увидеть ВСЕ меры одним текстом, вставлять/править много мер за раз.
>>     Пришли мне его содержимое (или фото по частям) — отдам готовый TMDL-скрипт оптимизации.
>>
>> C6. ПОРЯДОК, ЧТОБЫ НИЧЕГО НЕ СЛОМАТЬ
>>     копия .pbix -> одна страница -> сверка 3 цифр до/после -> старые меры в папку "_old"
>>     и скрыть -> через неделю удалить.


#####################################################################
 ПРОВЕРКА
#####################################################################
>> [ ] Ссылка ECB открывается в браузере на рабочем ПК
>> [ ] MarketData: 8 Series, страны IT/DE/ES..., Tenor 1Y..30Y
>> [ ] Spread vs Bund для Italy ~ значение, которое видишь в новостях (плюс-минус месячное среднее)
>> [ ] Слайсер страны меняет карточки и графики
>> [ ] Ползунок Shock Spread = 50 -> Impact Spread pp = Sens "+50 bps govies" компании
>> [ ] Solvency-страница в новом дизайне — скрин мне
>> [ ] Список мер / TMDL — мне для Части C
