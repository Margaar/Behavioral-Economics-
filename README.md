=====================================================================
 РЕАЛЬНЫЕ ИМЕНА КОМПАНИЙ ВМЕСТО "Company #1 / #2 ..."
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> ИДЕЯ
>> В каждом файле входные таблицы подписаны "Company #1, #2 ...". Кто есть кто — знает
>> только сам файл: на листе Local RAF справа "Company #1/#2/#3" стоят в тех же строках,
>> что ES003 / ES005 / ES006 слева.
>>   CompanyMap  = мостик "Company #n -> код" (из файла) + имя и страна из EntityList.
>>   FactData    = в конце подменяем Entity на "ES003 – Generali España ...".
>> Меры менять не нужно: слайсеры, матрицы, комментарии и Summary покажут новые имена сами.
>> Сделай копию .pbix перед началом.


#####################################################################
 ШАГ 1. ПРОВЕРЬ EntityList (1 минута)
#####################################################################
>> Trasforma dati > EntityList. Нужны колонки:  Entity (имя) | Code (ES003...) | Country.
>> Если колонки называются иначе — запомни имена, подставишь их в ШАГЕ 2 (строка X = ...).
>> Коды должны совпадать с колонкой "Legal Entity Code" на листе Local RAF (ES003, не "ES 003").


#####################################################################
 ШАГ 2. НОВЫЙ ЗАПРОС CompanyMap
#####################################################################
>> Home > Nuova origine > Query vuota > Editor avanzato > вставить всё > Fine.
>> Переименуй запрос в  CompanyMap. Правой кнопкой > Abilita caricamento: ВЫКЛ.

let
    Sheet = Workbook{[Item = "Local RAF", Kind = "Sheet"]}[Data],
    Rows  = List.Buffer(Table.ToRows(Table.ReplaceErrorValues(Sheet,
                List.Transform(Table.ColumnNames(Sheet), each {_, null})))),
    Txt   = (v) => if v = null then null else Text.Trim(Text.From(v)),
    Find  = (t as text) => List.First(List.Combine(List.Transform(List.Positions(Rows), (r) =>
                List.Transform(List.Select(List.Positions(Rows{r}), (c) =>
                    let x = Txt(Rows{r}{c}) in x <> null and Text.StartsWith(x, t)),
                    (c) => [R = r, C = c]))), null),
    cName = Find("Long Description")[C],
    cCode = Find("Legal Entity")[C],
    Recs  = List.Combine(List.Transform(List.Positions(Rows), (r) =>
                List.Transform(List.Select(List.Positions(Rows{r}), (c) =>
                    let x = Txt(Rows{r}{c}) in x <> null and Text.StartsWith(Text.Upper(x), "COMPANY #")),
                    (c) => [Company = Txt(Rows{r}{c}), Code = Txt(Rows{r}{cCode}), Name = Txt(Rows{r}{cName})]))),
    Ok    = List.Select(Recs, each [Code] <> null and [Code] <> "" and [Name] <> null and [Name] <> ""),
    T     = Table.FromRecords(Ok, type table [Company = text, Code = text, Name = text]),
    J     = Table.NestedJoin(T, {"Code"}, EntityList, {"Code"}, "E", JoinKind.LeftOuter),
    X     = Table.ExpandTableColumn(J, "E", {"Entity", "Country"}, {"EntName", "Country"}),
    Lbl   = Table.AddColumn(X, "Label", each [Code] & " – " &
                (if [EntName] <> null then [EntName] else [Name]), type text),
    Result = Table.Distinct(Lbl, {"Company"})
in
    Result

>> ПРОВЕРКА (dummy):
>>   Company #1 | ES003 | ... | Label "ES003 – Generali España De Seguros ..."
>>   Company #2 | ES005 | ... | "ES005 – Cajamar Vida ..."
>>   Company #3 | ES006 | ... | "ES006 – Cajamar Seguros Generales ..."
>> Если коды съехали на строку (Company #1 -> ES005) или строк 0 — пришли фото, поправлю.
>> Если в EntityList другие имена колонок: в строке X поменяй "Entity" / "Country" на свои.
>> Если EntName пустой — кода нет в EntityList, берётся имя из файла (ничего не ломается).


#####################################################################
 ШАГ 3. ПОДМЕНА Entity В КОНЦЕ FactData
#####################################################################
>> FactData > Editor avanzato. Внизу сейчас (после прошлых правок):
>>
>>     NoEntCountry = Table.ReplaceValue( ... {"Entity"})
>> in
>>     NoEntCountry
>>
>> 1) В конце строки NoEntCountry (после  {"Entity"})  ) поставь запятую.
>> 2) Замени две последние строки  "in / NoEntCountry"  на:

    Mapped = Table.NestedJoin(NoEntCountry, {"Entity"}, CompanyMap, {"Company"}, "M", JoinKind.LeftOuter),
    Exp    = Table.ExpandTableColumn(Mapped, "M", {"Label", "Country"}, {"EntLabel", "Country"}),
    Ent2   = Table.AddColumn(Exp, "Entity2", each if [EntLabel] <> null then [EntLabel] else [Entity], type text),
    Final2 = Table.RenameColumns(Table.RemoveColumns(Ent2, {"Entity", "EntLabel"}), {{"Entity2", "Entity"}})
in
    Final2

>> Если последний шаг у тебя называется не NoEntCountry — подставь его имя в строку Mapped.
>> Внизу окна: "Non sono stati rilevati errori di sintassi" > Fine > Chiudi e applica.
>> (Новая колонка Country пригодится, когда подключим все файлы стран.)


#####################################################################
 ШАГ 4. ПРОВЕРКА ПОСЛЕ ЗАГРУЗКИ
#####################################################################
>> [ ] Слайсер Entity: "ES003 – ...", "ES005 – ...", "ES006 – ...", "Company #4".
>> [ ] Solvency, SAA, FTC, Life/Non-Life KPIs: цифры по компании те же, что были
>>     (сравни одну цифру до/после, например Total Life SII VNP 27 у ES003).
>> [ ] Комментарии: вместо "Company #1:" пишут "ES003 – ...:".
>> [ ] Summary – Companies: тепловая карта со строками по новым именам.
>> [ ] Фильтры визуалов, где была выбрана конкретная "Company #1" (если были) —
>>     выбери компанию заново: старое значение больше не существует.
>> [ ] Страница-закладки/сохранённые выборы слайсеров — тоже перевыбрать.


#####################################################################
 ЧТО ОСТАЁТСЯ "Company #4"
#####################################################################
>> Холдинг (Company #4: Remittance Holding, HMS) не указан в Local RAF, поэтому
>> остаётся "Company #4". Если в файле есть лист со списком ВСЕХ компаний
>> (Start tab / Perimeter / Companies in scope с колонкой "Company #n" и кодом) —
>> пришли фото: переключу CompanyMap на этот лист, и холдинг тоже получит имя.
>> Временное решение для одного файла: в EntityList ничего не меняй, а в шаге
>> Ent2 замени  else [Entity]  на
>>     else if [Entity] = "Company #4" then "Holding – Company #4" else [Entity]
