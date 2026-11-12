=====================================================================
 CDD MARKET: ИСПРАВЛЕНИЕ МЕР С ОШИБКАМИ (⚠ в панели Dati)
 Строки с ">>" не копировать. Копировать только код.
 Как: клик по мере в панели Dati > строка формул > Ctrl+A > вставить > Enter.
=====================================================================

>> ПРИЧИНЫ (3 штуки, остальные ⚠ — "по цепочке" и уйдут сами)
>> 1) Scn Rates bp / Scn Spread bp / Scn Equity pct: в итальянском Power BI мера параметра
>>    называется не "[Shock Rates Valore]", а иначе ("Valore Shock Rates").
>>    -> берём значение прямо из таблицы параметра, без его меры.
>> 2) EuroStoxx Move %: меры [EuroStoxx Last] не существует (её надо было создать по шаблону).
>> 3) Curve Ref, Gov10Y M, HICP M: мера внутри фильтра CALCULATE ([Mkt Ref Date], [Mkt Country])
>>    -> ошибка "PLACEHOLDER". Мера должна сначала попасть в VAR.


#####################################################################
 ШАГ 1. ПАРАМЕТРЫ СЦЕНАРИЯ (3 меры)
#####################################################################
>> Проверка имён: в панели Dati есть таблицы "Shock Rates", "Shock Spread", "Shock Equity",
>> в каждой колонка с тем же именем. Если имя другое — поправь в формуле.

Scn Rates bp = SELECTEDVALUE ( 'Shock Rates'[Shock Rates], 0 )

Scn Spread bp = SELECTEDVALUE ( 'Shock Spread'[Shock Spread], 0 )

Scn Equity pct = SELECTEDVALUE ( 'Shock Equity'[Shock Equity], 0 )

>> Ползунки на фото стоят на 100 и 40 у левого края -> проверь параметры:
>> Modellazione > выдели таблицу параметра > значения: Spread -100..300, Rates -100..100,
>> Equity -40..20. Для старта поставь ползунки на 0.


#####################################################################
 ШАГ 2. EUROSTOXX (создать недостающую меру)
#####################################################################
>> Nuova misura:

EuroStoxx Last =
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "EUROSTOXX50" )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

>> Заодно проверь, что существуют (создай, если нет — тот же шаблон, меняется Series):

Euribor3M Last =
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "EURIBOR3M" )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

ECB DFR Last =
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "ECB_DFR" )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

EURUSD Last =
VAR t = CALCULATETABLE ( MarketData, REMOVEFILTERS ( MarketData ), MarketData[Series] = "EURUSD" )
VAR d = MAXX ( t, MarketData[Date] )
RETURN MAXX ( FILTER ( t, MarketData[Date] = d ), MarketData[Value] )

>> EuroStoxx Move % ПОСЛЕ этого заработает без изменений.
>> Если серия EUROSTOXX50 не загрузилась (в MarketData её нет) — мера будет пустой, ошибки нет.


#####################################################################
 ШАГ 3. МЕРЫ С МЕРОЙ ВНУТРИ ФИЛЬТРА (замени целиком)
#####################################################################

Curve Ref =
VAR rd = [Mkt Ref Date]
VAR d  = CALCULATE ( MAX ( MarketData[Date] ), REMOVEFILTERS ( MarketData ),
             MarketData[Series] = "EUR_AAA_CURVE", MarketData[Date] <= rd )
RETURN CALCULATE ( AVERAGE ( MarketData[Value] ),
           MarketData[Series] = "EUR_AAA_CURVE", MarketData[Date] = d )

Curve Last =
VAR d = CALCULATE ( MAX ( MarketData[Date] ), REMOVEFILTERS ( MarketData ),
            MarketData[Series] = "EUR_AAA_CURVE" )
RETURN CALCULATE ( AVERAGE ( MarketData[Value] ),
           MarketData[Series] = "EUR_AAA_CURVE", MarketData[Date] = d )

Gov10Y M =
VAR c = [Mkt Country]
RETURN CALCULATE ( AVERAGE ( MarketData[Value] ),
           MarketData[Series] = "GOV10Y", MarketData[Country] = c )

HICP M =
VAR c = [Mkt Country]
RETURN CALCULATE ( AVERAGE ( MarketData[Value] ),
           MarketData[Series] = "HICP", MarketData[Country] = c )

>> Gov10Y Last, HICP Last, Spread vs Bund Ref bp уже написаны через VAR — их не трогать.


#####################################################################
 ШАГ 4. ПРОВЕРКА
#####################################################################
>> [ ] В панели Dati (поиск "mark" / "scn") не осталось ⚠.
>>     Если ⚠ осталась на мере — открой её, пришли фото жёлтой строки.
>> [ ] Карточка Gov10Y Last при выбранной стране Italy ~ 3-4 %.
>> [ ] Spread vs Bund bp для Italy ~ 80-130 bp (зависит от месяца).
>> [ ] Ползунок Shock Spread = 50 -> Impact Spread pp у компании = её "+50 bps govies" из Sens.
>> [ ] Если Sens Govies pp / Sens IR Up pp пустые -> в CDD шок назван иначе:
>>     пришли список значений FactData[Sens Shock] — поправлю CONTAINSSTRING.
