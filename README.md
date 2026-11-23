=====================================================================
 FTC YE26 PRE / POST SOLVENCY II REVIEW — ИНСТРУКЦИЯ: POWER BI
 Источник: файлы CP-Tool 27-29, лист "Free Tangible Capital analysis"
 Pre  = колонка "Forecast 2026"
 Post = колонка "Forecast 2026 SII Review"
=====================================================================

---------------------------------------------------------------------
 ШАГ 0. ПОДГОТОВКА
---------------------------------------------------------------------
- Закрыть все файлы CP-Tool в Excel (открытый файл блокируется -> ошибка).
- Предупреждение "Microsoft ha bloccato le macro" игнорировать:
  Power Query макросы не запускает, файл читается нормально.
- Скопировать путь к файлу: Shift + правый клик -> "Copia come percorso".

---------------------------------------------------------------------
 ШАГ 1. ФУНКЦИЯ fnCPTool
---------------------------------------------------------------------
Home -> Trasforma dati -> Nuova origine -> Query vuota
-> Editor avanzato -> вставить код целиком -> Fine
-> переименовать query в: fnCPTool

(wb as binary, fileName as text) as table =>
let
    Labels  = {
        "Free Tangible Capital",
        "Eligible Own Funds",
        "Excess Capital",
        "Tangible Capital",
        "LGAAP Tangible Capital (LGAAP TCAP)",
        "Local GAAP Shareholders' Equity"
    },
    Src     = Excel.Workbook(wb, null, true),
    Sh      = Src{[Item = "Free Tangible Capital analysis", Kind = "Sheet"]}[Data],
    Raw     = Table.ToRows(Sh),
    Txt     = List.Transform(Raw, (r) => List.Transform(r, each if _ is text then Text.Lower(Text.Trim(_)) else null)),
    Idx     = List.Positions(Txt),

    // колонки по заголовкам
    HdrI    = List.First(List.Select(Idx, each List.Contains(Txt{_}, "forecast 2026 sii review"))),
    Hdr     = Txt{HdrI},
    cPre    = List.PositionOf(Hdr, "forecast 2026"),
    cPost   = List.PositionOf(Hdr, "forecast 2026 sii review"),

    // IFICC из ячейки выбора; если пусто -> имя файла
    IfI     = List.First(List.Select(Idx, each List.AnyTrue(List.Transform(Txt{_}, (c) => c <> null and Text.StartsWith(c, "please select the company"))))),
    IfVal   = if IfI = null then null else
                List.First(List.Select(Raw{IfI}, each _ <> null and not Text.StartsWith(Text.Lower(Text.From(_)), "please select"))),
    IFICC   = if IfVal = null then fileName else Text.Trim(Text.From(IfVal)),

    // нужные строки
    Out     = List.Transform(Labels, (L) =>
                let i = List.First(List.Select(Idx, each List.Contains(Txt{_}, Text.Lower(L))))
                in  [ IFICC = IFICC,
                      File  = fileName,
                      Item  = L,
                      Pre   = if i = null then null else Raw{i}{cPre},
                      Post  = if i = null then null else Raw{i}{cPost} ]),
    Tbl     = Table.FromRecords(Out),
    Typed   = Table.TransformColumnTypes(Tbl, {{"Pre", type number}, {"Post", type number}, {"Item", type text}, {"IFICC", type text}})
in
    Typed

---------------------------------------------------------------------
 ШАГ 2. ТЕСТ НА ОДНОМ ФАЙЛЕ
---------------------------------------------------------------------
Nuova origine -> Query vuota -> в строку формул вставить (подставить путь):

= fnCPTool(File.Contents("C:\...\CP-Tool 27-29_MASTER_v2.3.xlsm"), "MASTER")

Переименовать в TEST.
Проверка (MASTER):
  Free Tangible Capital   Pre 19.822.796,7   Post 21.621.007,5
  Eligible Own Funds      Pre 50.972.905,7   Post 50.590.608,9
Совпало -> функция работает. Query TEST потом можно удалить.

---------------------------------------------------------------------
 ШАГ 3. ПАРАМЕТР С ПАПКОЙ
---------------------------------------------------------------------
Home -> Gestisci parametri -> Nuovo parametro
  Nome:   FolderPath
  Tipo:   Testo
  Valore: C:\...\папка_с_CP-Tool   (без "\" в конце)

---------------------------------------------------------------------
 ШАГ 4. QUERY CP_Data (все файлы из папки)
---------------------------------------------------------------------
Query vuota -> Editor avanzato -> вставить -> переименовать в CP_Data

let
    Files   = Folder.Files(FolderPath),
    Xl      = Table.SelectRows(Files, each
                List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
                and not Text.StartsWith([Name], "~$")),
    Res     = Table.AddColumn(Xl, "Res", each try fnCPTool([Content], [Name])),
    Status  = Table.AddColumn(Res, "Status", each if [Res][HasError] then [Res][Error][Message] else "OK", type text),
    OK      = Table.SelectRows(Status, each [Status] = "OK"),
    Data    = Table.AddColumn(OK, "T", each [Res][Value]),
    Comb    = Table.Combine(Data[T])
in
    Comb

Если файлы на SharePoint — заменить первую строку:
    Files = Table.SelectRows(
              SharePoint.Files("https://xxx.sharepoint.com/sites/yyy", [ApiVersion = 15]),
              each Text.Contains([Folder Path], "CP-Tool")),

---------------------------------------------------------------------
 ШАГ 5. QUERY CP_Errors (какие файлы не прочитались)
---------------------------------------------------------------------
let
    Files   = Folder.Files(FolderPath),
    Xl      = Table.SelectRows(Files, each
                List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
                and not Text.StartsWith([Name], "~$")),
    Res     = Table.AddColumn(Xl, "Res", each try fnCPTool([Content], [Name])),
    Status  = Table.AddColumn(Res, "Status", each if [Res][HasError] then [Res][Error][Message] else "OK", type text),
    Bad     = Table.SelectRows(Status, each [Status] <> "OK"),
    Keep    = Table.SelectColumns(Bad, {"Name", "Status"})
in
    Keep

-> Chiudi e applica

---------------------------------------------------------------------
 ШАГ 6. МЕРЫ DAX (в €m), папка Free Tangible Capital\S2 Review
---------------------------------------------------------------------
FTC Pre =
DIVIDE ( CALCULATE ( SUM ( CP_Data[Pre] ),  CP_Data[Item] = "Free Tangible Capital" ), 1e6 )

FTC Post =
DIVIDE ( CALCULATE ( SUM ( CP_Data[Post] ), CP_Data[Item] = "Free Tangible Capital" ), 1e6 )

FTC Δ =
[FTC Post] - [FTC Pre]

FTC Δ % =
DIVIDE ( [FTC Δ], ABS ( [FTC Pre] ) )

Формат: суммы #,##0.0 ; проценты 0.0%

---------------------------------------------------------------------
 ШАГ 7. (ОПЦИОНАЛЬНО) СПИСОК КОМПАНИЙ S2 ИЗ SCOPE
---------------------------------------------------------------------
Recupera dati -> Excel -> "2026_Payout Ratio & FTC Analysis" -> лист Scope
В Power Query:
  - Rimuovi righe superiori: 6
  - Usa la prima riga come intestazione
  - Оставить колонки: Code, Long Description, SII Projection
  - Фильтр SII Projection = "X"
Связь в модели: Scope[Code]  1 —> *  CP_Data[IFICC]
-> видно названия компаний и тех S2, у кого нет файла.

---------------------------------------------------------------------
 ШАГ 8. ВИЗУАЛ
---------------------------------------------------------------------
Tabella: IFICC (или Scope[Long Description]) | FTC Pre | FTC Post | FTC Δ | FTC Δ %
Formato -> Totali -> ON
Для Marco: "..." на визуале -> Esporta dati -> Excel

---------------------------------------------------------------------
 ЕСЛИ ЧТО-ТО НЕ РАБОТАЕТ
---------------------------------------------------------------------
- "Key didn't match any rows"   -> в файле нет листа "Free Tangible Capital analysis"
                                   (другое имя листа / другой шаблон)
- Pre/Post пустые               -> заголовок колонки отличается от
                                   "Forecast 2026" / "Forecast 2026 SII Review"
- IFICC = имя файла             -> код IFICC не найден в ячейке выбора;
                                   норм, если в имени файла есть код компании
- "file in use"                 -> закрыть файл в Excel
- Обновление после новых файлов -> просто Aggiorna
=====================================================================



=====================================================================
 FTC YE26 PRE / POST SOLVENCY II REVIEW — ИНСТРУКЦИЯ: EXCEL POWER QUERY
 (вариант без Power BI: результат — готовая таблица в Excel)
 Источник: файлы CP-Tool 27-29, лист "Free Tangible Capital analysis"
 Pre  = колонка "Forecast 2026"
 Post = колонка "Forecast 2026 SII Review"
=====================================================================

---------------------------------------------------------------------
 ШАГ 0. ПОДГОТОВКА
---------------------------------------------------------------------
- Создать НОВЫЙ пустой файл Excel, например FTC_S2Review.xlsx
  (не работать внутри CP-Tool).
- Закрыть все файлы CP-Tool (открытый файл блокируется -> ошибка).
- Путь к файлу/папке: Shift + правый клик -> "Copia come percorso".

---------------------------------------------------------------------
 ШАГ 1. ФУНКЦИЯ fnCPTool
---------------------------------------------------------------------
Dati -> Recupera dati -> Da altre origini -> Query vuota
-> откроется Editor di Power Query
-> Home -> Editor avanzato -> вставить код целиком -> Fine
-> справа в "Proprietà" переименовать в: fnCPTool

(wb as binary, fileName as text) as table =>
let
    Labels  = {
        "Free Tangible Capital",
        "Eligible Own Funds",
        "Excess Capital",
        "Tangible Capital",
        "LGAAP Tangible Capital (LGAAP TCAP)",
        "Local GAAP Shareholders' Equity"
    },
    Src     = Excel.Workbook(wb, null, true),
    Sh      = Src{[Item = "Free Tangible Capital analysis", Kind = "Sheet"]}[Data],
    Raw     = Table.ToRows(Sh),
    Txt     = List.Transform(Raw, (r) => List.Transform(r, each if _ is text then Text.Lower(Text.Trim(_)) else null)),
    Idx     = List.Positions(Txt),

    HdrI    = List.First(List.Select(Idx, each List.Contains(Txt{_}, "forecast 2026 sii review"))),
    Hdr     = Txt{HdrI},
    cPre    = List.PositionOf(Hdr, "forecast 2026"),
    cPost   = List.PositionOf(Hdr, "forecast 2026 sii review"),

    IfI     = List.First(List.Select(Idx, each List.AnyTrue(List.Transform(Txt{_}, (c) => c <> null and Text.StartsWith(c, "please select the company"))))),
    IfVal   = if IfI = null then null else
                List.First(List.Select(Raw{IfI}, each _ <> null and not Text.StartsWith(Text.Lower(Text.From(_)), "please select"))),
    IFICC   = if IfVal = null then fileName else Text.Trim(Text.From(IfVal)),

    Out     = List.Transform(Labels, (L) =>
                let i = List.First(List.Select(Idx, each List.Contains(Txt{_}, Text.Lower(L))))
                in  [ IFICC = IFICC,
                      File  = fileName,
                      Item  = L,
                      Pre   = if i = null then null else Raw{i}{cPre},
                      Post  = if i = null then null else Raw{i}{cPost} ]),
    Tbl     = Table.FromRecords(Out),
    Typed   = Table.TransformColumnTypes(Tbl, {{"Pre", type number}, {"Post", type number}, {"Item", type text}, {"IFICC", type text}})
in
    Typed

---------------------------------------------------------------------
 ШАГ 2. ТЕСТ НА ОДНОМ ФАЙЛЕ
---------------------------------------------------------------------
В Editor di Power Query: Home -> Nuova origine -> Altre origini -> Query vuota
В строку формул (подставить путь):

= fnCPTool(File.Contents("C:\...\CP-Tool 27-29_MASTER_v2.3.xlsm"), "MASTER")

Проверка (MASTER):
  Free Tangible Capital   Pre 19.822.796,7   Post 21.621.007,5
  Eligible Own Funds      Pre 50.972.905,7   Post 50.590.608,9
Совпало -> удалить эту тестовую query.

---------------------------------------------------------------------
 ШАГ 3. ПАРАМЕТР С ПАПКОЙ
---------------------------------------------------------------------
Home -> Gestisci parametri -> Nuovo parametro
  Nome:   FolderPath
  Tipo:   Testo
  Valore: C:\...\папка_с_CP-Tool

---------------------------------------------------------------------
 ШАГ 4. QUERY CP_Data (длинная таблица, все показатели)
---------------------------------------------------------------------
Query vuota -> Editor avanzato -> вставить -> переименовать в CP_Data

let
    Files   = Folder.Files(FolderPath),
    Xl      = Table.SelectRows(Files, each
                List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
                and not Text.StartsWith([Name], "~$")),
    Res     = Table.AddColumn(Xl, "Res", each try fnCPTool([Content], [Name])),
    Status  = Table.AddColumn(Res, "Status", each if [Res][HasError] then [Res][Error][Message] else "OK", type text),
    OK      = Table.SelectRows(Status, each [Status] = "OK"),
    Data    = Table.AddColumn(OK, "T", each [Res][Value]),
    Comb    = Table.Combine(Data[T])
in
    Comb

SharePoint вместо локальной папки — заменить первую строку:
    Files = Table.SelectRows(
              SharePoint.Files("https://xxx.sharepoint.com/sites/yyy", [ApiVersion = 15]),
              each Text.Contains([Folder Path], "CP-Tool")),

---------------------------------------------------------------------
 ШАГ 5. QUERY FTC_Table (итоговая таблица для Christian, в €m)
---------------------------------------------------------------------
Query vuota -> Editor avanzato -> вставить -> переименовать в FTC_Table

let
    Src     = CP_Data,
    FTC     = Table.SelectRows(Src, each [Item] = "Free Tangible Capital"),
    Keep    = Table.SelectColumns(FTC, {"IFICC", "File", "Pre", "Post"}),
    Mln     = Table.TransformColumns(Keep, {
                {"Pre",  each if _ = null then null else _ / 1e6, type number},
                {"Post", each if _ = null then null else _ / 1e6, type number}}),
    Delta   = Table.AddColumn(Mln, "Delta", each
                if [Pre] = null or [Post] = null then null else [Post] - [Pre], type number),
    DeltaPc = Table.AddColumn(Delta, "Delta %", each
                if [Pre] = null or [Pre] = 0 or [Delta] = null then null else [Delta] / Number.Abs([Pre]), Percentage.Type),
    Ren     = Table.RenameColumns(DeltaPc, {
                {"Pre",  "FTC YE26 pre S2 rev (€m)"},
                {"Post", "FTC YE26 post S2 rev (€m)"},
                {"Delta", "Δ (€m)"}}),
    Sorted  = Table.Sort(Ren, {{"IFICC", Order.Ascending}})
in
    Sorted

---------------------------------------------------------------------
 ШАГ 6. QUERY CP_Errors (какие файлы не прочитались)
---------------------------------------------------------------------
let
    Files   = Folder.Files(FolderPath),
    Xl      = Table.SelectRows(Files, each
                List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
                and not Text.StartsWith([Name], "~$")),
    Res     = Table.AddColumn(Xl, "Res", each try fnCPTool([Content], [Name])),
    Status  = Table.AddColumn(Res, "Status", each if [Res][HasError] then [Res][Error][Message] else "OK", type text),
    Bad     = Table.SelectRows(Status, each [Status] <> "OK"),
    Keep    = Table.SelectColumns(Bad, {"Name", "Status"})
in
    Keep

---------------------------------------------------------------------
 ШАГ 7. ЗАГРУЗКА В EXCEL
---------------------------------------------------------------------
Home -> Chiudi e carica in...
  - FTC_Table   -> Tabella, nuovo foglio      (это отдаём Christian / Marco)
  - CP_Data     -> Tabella, nuovo foglio      (все показатели, для сверки)
  - CP_Errors   -> Tabella, nuovo foglio      (контроль)
  - fnCPTool    -> Crea solo connessione
Если Excel загрузил всё сразу: Dati -> Query e connessioni ->
правый клик на query -> Carica in... -> выбрать нужное.

Форматирование на листе FTC_Table:
  - суммы: #.##0,0
  - Δ %:   0,0%
  - Итоговая строка: клик в таблицу -> Progettazione tabella -> Riga totale -> ON
    (в итоговой строке для Δ % поставить формулу = Δ_total / ABS(Pre_total),
     а не сумму процентов)

---------------------------------------------------------------------
 ШАГ 8. (ОПЦИОНАЛЬНО) ТОЛЬКО КОМПАНИИ S2 + НАЗВАНИЯ
---------------------------------------------------------------------
Dati -> Recupera dati -> Da file -> Da cartella di lavoro ->
"2026_Payout Ratio & FTC Analysis" -> лист Scope -> Trasforma dati:
  - Rimuovi righe superiori: 6
  - Usa la prima riga come intestazione
  - Оставить: Code, Long Description, SII Projection
  - Фильтр SII Projection = "X"
  - Переименовать query в Scope_S2

В FTC_Table перед строкой "Sorted" добавить объединение
(Home -> Unisci query -> FTC_Table.IFICC = Scope_S2.Code,
 Tipo: Esterno sinistro) -> развернуть Long Description.
Чтобы увидеть компании S2 БЕЗ файла: тип объединения "Esterno destro"
на отдельной копии query.

---------------------------------------------------------------------
 ОБНОВЛЕНИЕ
---------------------------------------------------------------------
Новые файлы в папке -> Dati -> Aggiorna tutto. Больше ничего делать не нужно.

---------------------------------------------------------------------
 ЕСЛИ ЧТО-ТО НЕ РАБОТАЕТ
---------------------------------------------------------------------
- "Key didn't match any rows"   -> в файле нет листа "Free Tangible Capital analysis"
- Pre/Post пустые               -> заголовок колонки отличается от
                                   "Forecast 2026" / "Forecast 2026 SII Review"
- IFICC = имя файла             -> код IFICC не найден в ячейке выбора
- "file in use"                 -> закрыть файл в Excel
- Formula.Firewall / privacy    -> File -> Opzioni e impostazioni -> Opzioni query ->
                                   Privacy -> "Ignora i livelli di privacy" (для этого файла)
=====================================================================

