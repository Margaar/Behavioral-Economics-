
=====================================================================


0. 
---------------------------------------------------------------------
Строка 39 "Total End of Period (EoP)": в колонке кода стоит OPN_ADJ_CUST_ID1
(это код Opening Restatement #1). Заменить на EOP.
EoP НЕ тянуть по коду, а считать формулой (см. п.4).


1. СООТВЕТСТВИЕ КОЛОНОК
---------------------------------------------------------------------
CP-Tool                          ->  CPR
A  (код строки)                  ->  C  (Account Code) = КЛЮЧ ДЛЯ ФОРМУЛ
B  = ROW  (строка ввода)         ->  тянуть данные по коду
B  = CLM  (подытог, формула)     ->  считать формулой SUM в CPR, не тянуть
F  (название строки)             ->  E
G  FY25 Actual                   ->  F  Actual 2025   (из CP-CoRE)
H  FY26 Forecast                 ->  G  Forecast 2026 (из CP-CoRE)
I-K Plan 2027-2029               ->  НЕ ПЕРЕНОСИТЬ
-                                ->  H  Actual 2026 = из ручной вставки в J (Tagetik)

* ROW/CLM = гипотеза. Проверить на 2-3 строках: у CLM в ячейке формула суммы.


2. ЧТО ПЕРЕНОСИТЬ (порядок строк = как в источнике)
---------------------------------------------------------------------
Блок "Before Tax AoM items":

Total Beginning of Period (BoP)                      BOP                      данные
Opening Restatement                                  подытог                  =SUM(5 строк)
   Opening Restatement - Item #1..#5                 OPN_ADJ_CUST_ID1..5      данные
Regulatory and model changes                         подытог                  = EIOPA + non-EIOPA + Model
   EIOPA changes                                     подытог
      Volatility Adjustment reference portfolio      SII_OFM_REG_EIOPA_PTF
      Ultimate Forward Rate                          SII_OFM_REG_EIOPA_UFR
      Risk Free reference yield curves               SII_OFM_REG_EIOPA_YLD
      Transitional measures TP & Other               REG_EIOPA_TMREL
      SII Review impact                              подытог
         Risk Free Rate & Volatility Adjustment      SII_OFM_REG_SII_REV_RFR
         Risk Margin                                 SII_OFM_REG_SII_REV_RM
         Delta Net DTL                               SII_OFM_REG_SII_REV_NDT
         Participations impact                       REG_SII_REV_PART
         Other Impacts                               REG_SII_REV_OF
   Regulatory changes other than EIOPA               подытог = Life + Non-Life + Other
      Life      custom #1..#3                        SII_OFM_REG_LMCHG_CUST_ID1..3
      Non-Life  custom #1..#3                        SII_OFM_REG_NMCHG_CUST_ID1..3
      Other     custom #1..#3                        SII_OFM_REG_OMCHG_CUST_ID1..3
   Model enhancements & Other                        подытог = Life + Non-Life + Other
      Life      custom #1..#3                        SII_OFM_REG_LOTH_CUST_ID1..3
      Non-Life  custom #1..#3                        SII_OFM_REG_NOTH_CUST_ID1..3
      Other     custom #1..#3                        SII_OFM_REG_OOTH_CUST_ID1..3
Normalised Capital Generation                        подытог
   Life / Non-Life / Holding / Financials            подытоги + строки SII_OFM_CPGEN_...
Economic variances                                   подытог + IR, Spread & VA, Equity,
                                                     Property, Volatilities, Inflation, Other
Non-Economic variances -> Operating variances (Life / Non-Life) -> Other Operating
Other Variances (Life PVFP2, Expenses, Other Delta OF adjustments)
Residual unallocated variances
Tax impact
M&A and portfolio transfers
Capital movements (injections, sub debt, dividends, foreseeable dividend)
Total End of Period (EoP)                            EOP                      ФОРМУЛА

CUSTOM ITEMS (<...Item #1>, CUST_ID):
- у некоторых компаний в них реальные цифры -> НЕ удалять
- переносить, но в колонке A ставить HIDE, если сумма по строке = 0
- сворачивать группировкой

НЕ ПЕРЕНОСИТЬ:
- Plan 2027-2029
- IFRS/MVBS carve-out строки
- SCR-часть (колонки M-Q), пока тутор не скажет иначе


3. ФОРМУЛЫ
---------------------------------------------------------------------
Строки с данными (ROW):
F:  =SUMIFS(Src[Value]; Src[Code]; $C18; Src[Entity]; $Entity; Src[Period]; "FY25")
G:  то же самое с "FY26"
H:  =J18*$K$5     (Actual 2026 из вставки Tagetik x reporting scale)

Src = лист с данными CP-CoRE (Input или 7.4 Data Export - уточнить).
H = J x K5 = гипотеза (K5 "Reporting scale in OF_AOM file").

Подытоги (CLM):
=SUM(дочерние строки) или явная сумма как в источнике.
Пример из CP-Tool: Regulatory and model changes = G18+G29+G42
(EIOPA changes + Regulatory other than EIOPA + Model enhancements & Other)


4. КОНТРОЛИ
---------------------------------------------------------------------
[1] EoP расчётный:
    BoP + Opening + Reg&Model + CapGen + EcoVar + OpVar + OtherVar
    + Residual + Tax + M&A + CapMov
[2] EoP расчётный - EoP из CP-Tool                       = 0
[3] EoP Actual 2025 = BoP Forecast 2026                   разница = 0
    (в CP-Tool BoP FY26 = 67.424.478,4)
[4] Residual unallocated: доля в общем движении           должна быть малой
    Если в Residual вся сумма (пример: 85 млн в FY25) -> AoM не разнесён -> флаг


5. ИТОГОВЫЙ ВИД ЛИСТА
---------------------------------------------------------------------
A      | C (Code)             | E (Item)                         | F Act25 | G Fcst26 | H Act26 | Δ abs | Δ % | Check
-------+----------------------+----------------------------------+---------+----------+---------+-------+-----+------
       |                      | OWN FUNDS CAPITAL GENERATION     |         |          |         |       |     |
       | SII_OFM_CPGEN_L_VNP… | Life SII VNP (gross of tax)      |         |          |         |       |     |
       | …                    | NBV before RM / NBV RM           |         |          |         |       |     |
       | …NL_CYBER / NL_DISC  | Non-life CY tech result / Disc.  |         |          |         |       |     |
       |                      | OF ANALYSIS OF MOVEMENT (pre-tax)|         |          |         |       |     |
       | BOP                  | Total BoP                        |         |          |         |       |     |
HIDE   | OPN_ADJ_CUST_ID1     |   Opening Restatement #1         |         |          |         |       |     |
       | …                    | … (порядок из п.2)               |         |          |         |       |     |
       | EOP                  | Total EoP                        |         |          |         |       |     | OK/X

Справа (колонка J): зона ручной вставки Actual 2026 из Tagetik - оставить.
Δ abs = H - G ; Δ % = (H - G) / ABS(G)
Колонки Δ = моё предложение (без них это копия плана, а не performance review).


6. СПРОСИТЬ У ТУТОРА
---------------------------------------------------------------------
1. Нужна ли SCR-часть AoM или только Own Funds?
2. Из какого листа CPR брать данные CP-CoRE (Input / 7.4 Data Export)?
3. Custom items: оставлять с HIDE или сворачивать в подытог?
4. Добавлять колонки Δ и %?


ПОРЯДОК РАБОТЫ
---------------------------------------------------------------------
[ ] 1. Исправить код в строке EoP (п.0)
[ ] 2. Достроить строки по п.2 (коды в C, названия в E)
[ ] 3. Пометить подытоги, проставить SUM-формулы
[ ] 4. Проставить SUMIFS в строки с данными (F, G), H = J x K5
[ ] 5. Добавить контроли (п.4)
[ ] 6. Проставить HIDE для пустых custom items, сгруппировать
[ ] 7. (после ответа тутора) Δ-колонки, SCR-блок
