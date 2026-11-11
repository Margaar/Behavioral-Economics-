==================================================================


#####################################################################
 ШАГ 1. "KEY 3" ВМЕСТО НАЗВАНИЯ БЛОКА  (Power Query, FactData)
#####################################################################
>> Почему: на листе Non-Life Translation Framework слева от заголовка с периодами стоит
>> служебная метка "KEY 3" (скорее всего в скрытой колонке). Парсер взял её как название
>> блока — у обеих таблиц (итог и по компании) одно и то же имя.
>>
>> Trasforma dati > FactData > Editor avanzato.
>> 1) Найди строку, которая начинается с   WithKey  = Table.AddColumn(WithSrc, "LineKey", ...
>> 2) ПЕРЕД ней вставь:

    FixBlk   = Table.ReplaceValue(WithSrc, each [Block], each
                   if [Block] = null or Text.StartsWith(Text.Upper([Block]), "KEY")
                   then (if [Topic] = "Non-Life TF by Company" then "LGAAP to IFRS17 bridge – by company"
                         else if [Topic] = "Non-Life between LGAAP and IFRS" then "LGAAP to IFRS17 bridge – total"
                         else [Topic])
                   else [Block],
                   Replacer.ReplaceValue, {"Block"}),

>> 3) В строке WithKey замени   Table.AddColumn(WithSrc,   на   Table.AddColumn(FixBlk,
>> 4) Внизу окна: "Non sono stati rilevati errori di sintassi" > Fine > Chiudi e applica.
>>
>> Результат: два блока "LGAAP to IFRS17 bridge – total" и "... – by company".
>> Любое другое "KEY..." на других листах тоже заменится на название Topic.
>> Тексты в кавычках можно поменять на свои; условия [Topic] = "..." не трогать.


#####################################################################
 ШАГ 2. ПУСТЫЕ СТРАНИЦЫ (слайсер Entity синхронизирован)
#####################################################################
>> Почему: на BU-листах нет "Company #" (Entity пустой), а слайсер Entity
>> синхронизирован со всеми страницами и держит выбор Company #1 -> всё отфильтровано.
>>
>> 1) Visualizza > Sincronizza filtri dei dati (откроется панель).
>> 2) Выдели слайсер Entity на любой странице.
>> 3) В панели СНИМИ обе галочки (Sincronizza и Visibile) у страниц:
>>      Total Distributions BU, Local GAAP Details, Holding LGAAP,
>>      Profitability Ratios BU, Own Funds / OF Capital Generation Items,
>>      Non-Life Translation Framework*
>> 4) На этих страницах удали сам слайсер Entity (если остался).
>>
>> * Non-Life Translation Framework: правая таблица "by company" имеет Company #.
>>   Если хочешь выбирать компанию — оставь слайсер, НО:
>>   слайсер > Formato > Modifica interazioni > на ЛЕВОЙ матрице (total) значок "Nessuna".
>>
>> Проверка: Own Funds и Profitability Ratios BU должны показать строки.
>> Если Own Funds всё ещё пустой — пришли фото: возможно, у этого листа есть Company #.


#####################################################################
 ШАГ 3. ФИЛЬТРЫ TOPIC НА СТРАНИЦАХ
#####################################################################
>> Filtri in questa pagina > Topic. Сверь с таблицей:
>>
>>  Вкладка                          Topic (Filtri in questa pagina)
>>  -------------------------------  -----------------------------------------------------
>>  Total Distributions BU           Total Distributions BU
>>  Solvency II Entity               Solvency II
>>  Local Solvency Entity            Local Solvency
>>  SST Entity                       SST
>>  Holding Entity                   Holding
>>  Other Sect. Solvency Entity      Other Sectoral
>>  Local GAAP Details               Local GAAP Total + Local GAAP Life + Local GAAP Non-Life
>>  IFRS Details                     IFRS
>>  Holding LGAAP                    Local LGAAP              <- было Local GAAP Total
>>  Non-Life Translation Framework   Non-Life between LGAAP and IFRS + Non-Life TF by Company
>>  Profitability Ratios BU          BU Profitability
>>  Profitability Ratios LEs         Profitability Ratios
>>  ORMT Impacts                     ORMT
>>  Solvency AoM                     Solvency II AoM
>>  OF Capital Generation Items      Own Funds
>>  Sensitivities                    Sensitivities
>>
>> ДВЕ ТАБЛИЦЫ НА ОДНОЙ СТРАНИЦЕ:
>>  - Local GAAP Details: в Righe поставь FactData[Topic] > Block > Line
>>    (Total / Life / Non-Life будут отдельными группами).
>>  - Non-Life Translation Framework: ЛИБО Righe = Topic > Block > Line в одной матрице,
>>    ЛИБО две матрицы с фильтром визуала (Filtri in questo oggetto visivo > Topic):
>>      левая  = Non-Life between LGAAP and IFRS
>>      правая = Non-Life TF by Company
>>
>> ORMT Impacts: Righe = FactData[Section] > FactData[Line]  (группы <ORMT #1>, <ORMT #2>).


#####################################################################
 ШАГ 4. ПОРЯДОК КОЛОНОК И СТРОК КАК В EXCEL
#####################################################################
>> Сейчас: Actual 2025 | Actual 2026 | Forecast 2026 и строки по алфавиту (на Solvency II
>> даже обратном: Tier III, Tier II ...). Нужна сортировка по колонке.
>>
>> 4a) Если колонки PeriodSort нет: Visualizzazione dati > FactData > Nuova colonna:

PeriodSort =
FactData[Year] * 10
    + SWITCH ( FactData[PeriodType], "Forecast", 1, "Half-Year", 2, "Actual", 3, "Plan", 6, 9 )

>> 4b) Visualizzazione dati > FactData > клик по заголовку колонки >
>>     Strumenti colonna > Ordina per colonna:
>>        Period -> PeriodSort
>>        Line   -> LineOrder
>>        Block  -> BlockOrder
>>     Ошибка "più valori per lo stesso valore" -> пришли фото, какая колонка.
>>
>> 4c) В КАЖДОЙ матрице: "..." (вверху визуала) > Ordina asse > Line > Crescente.
>>     (если стоит сортировка по Block — тоже Crescente)


#####################################################################
 ШАГ 5. BU Line (если ещё не заменила)
#####################################################################
>> На старых фото в BU Line ещё "VAR rows" — это зарезервированное слово, мера в ошибке,
>> из-за неё падают BU Executive Summary и Summary BU. Замени формулу BU Line целиком:

BU Line =
VAR tR = ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] ), "@s", [CPR Status] )
VAR n  = COUNTROWS ( FILTER ( tR, [@s] IN { 1, 2, 3 } ) ) + 0
VAR g  = COUNTROWS ( FILTER ( tR, [@s] = 1 ) ) + 0
VAR rl = CONCATENATEX ( FILTER ( tR, [@s] = 3 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR al = CONCATENATEX ( FILTER ( tR, [@s] = 2 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR sr = CALCULATE ( [Key Act CY], SummaryIndex[Topic] = "Solvency II" )
RETURN IF ( n > 0,
    IF ( NOT ISBLANK ( sr ), "SII ratio " & FORMAT ( sr, "0%" ) & " · ", "" )
    & g & " of " & n & " in line"
    & IF ( rl <> "", " · Action: " & rl, "" )
    & IF ( al <> "", " · Explain: " & al, "" ) )


#####################################################################
 ШАГ 6. ДИЗАЙН (по желанию, но сильно улучшает вид)
#####################################################################
>> 6a) Заголовки страниц: убрать "Executive Summary" у SST / Holding / Other Sectoral
>>     (это страницы листов). Писать как в Excel: "SST Entity", "Holding Entity" ...
>>
>> 6b) Раскладка (страница 1280 x 720, одинаково на всех страницах):
>>
>>   ┌──────────────────────────────────────────────────────────────────────┐
>>   │ ШАПКА 0,0  1280x56 #7A0F0C: название листа (белый, 18) + CPR Subtitle  │
>>   │                      справа слайсер Entity (Elenco a discesa)          │
>>   ├───────────────────────────────────────────┬──────────────┬───────────┤
>>   │ МАТРИЦА "как в Excel"                       │ ОТКЛОНЕНИЯ    │ KEY       │
>>   │ X 16, Y 72, W 640, H 632                    │ X 664 W 300   │ MESSAGES  │
>>   │ Block > Line | Period | CPR Cell             │ Δ vs Fcst     │ X 972     │
>>   │                                              │ Var vs Fcst   │ W 292     │
>>   │                                              │ Var vs PY     │ CPR Comm. │
>>   └───────────────────────────────────────────┴──────────────┴───────────┘
>>
>>   - Матрица отклонений: убери Act CY и Act PY (они уже есть слева).
>>   - Key messages: белый прямоугольник + красная полоса слева 4 px #C21B17 +
>>     заголовок "Key messages" (Segoe UI Semibold 11, #C21B17) + карточка CPR Comment.
>>   - Слайсер Entity: Stile Elenco a discesa, без синей заливки
>>     (Formato > Valori > Sfondo: белый; Intestazione: "Entity", #A50E0E).
>>
>> 6c) Матрицы:
>>   - Шрифт: Valori и Intestazioni riga 10 (сейчас ~8, мелко).
>>   - Intestazioni riga: Block — Segoe UI Semibold, #A50E0E.
>>   - Строки Total/Ratio: в мере Cell BG поменяй "#F2E3E2" на "#EED5D3" (ярче),
>>     в Cell Font для k = 1 оставь "#7A0F0C".
>>   - Обе матрицы: одинаковая Spaziatura righe (Formato > Griglia), тогда строки
>>     слева и справа стоят на одном уровне.
>>
>> 6d) Шаблонные строки (<Other driver #1>, [custom to specify] ...):
>>     пока данные dummy — оставить. Когда будут реальные данные: в Righe у Line
>>     выключи "Mostra elementi senza dati" — останутся только заполненные строки.
>>
>> 6e) Сделай ОДНУ страницу по 6b–6c, потом скопируй шапку, матрицы и Key messages
>>     на остальные (Ctrl+C / Ctrl+V встают на те же координаты) и поменяй только
>>     фильтр Topic и название.


#####################################################################
 ПРОВЕРКА
#####################################################################
>> [ ] Non-Life Translation Framework: два блока вместо "KEY 3"
>> [ ] Local GAAP Details, Profitability Ratios BU, Own Funds, Holding LGAAP показывают строки
>> [ ] Holding LGAAP: строки "Revenue from participations / Expenses / Operating result"
>> [ ] Колонки: Actual 2025 | Forecast 2026 | Actual 2026 | Plan 2027
>> [ ] Строки и блоки в порядке Excel на всех страницах
>> [ ] BU Line без ошибки, Summary BU работает
>> [ ] Пришли 2–3 скриншота
