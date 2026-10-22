{
  "name": "Generali CDD",
  "dataColors": ["#C21B17", "#5C5C5C", "#E8746F", "#A8A8A8", "#7A0F0C", "#D9B26F", "#3F6E8C", "#8FB3C9"],
  "background": "#FFFFFF",
  "foreground": "#333333",
  "tableAccent": "#C21B17",
  "good": "#2E7D32",
  "neutral": "#E0A100",
  "bad": "#A50E0E",
  "maximum": "#C21B17",
  "center": "#F2F2F2",
  "minimum": "#5C5C5C",
  "hyperlink": "#C21B17",
  "visitedHyperlink": "#7A0F0C",
  "textClasses": {
    "callout": { "fontSize": 22, "fontFace": "Segoe UI Semibold", "color": "#C21B17" },
    "title":   { "fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#333333" },
    "header":  { "fontSize": 11, "fontFace": "Segoe UI Semibold", "color": "#333333" },
    "label":   { "fontSize": 9,  "fontFace": "Segoe UI",          "color": "#333333" }
  },
  "visualStyles": {
    "*": {
      "*": {
        "title": [{ "show": true, "fontColor": { "solid": { "color": "#333333" } }, "fontSize": 11, "fontFamily": "Segoe UI Semibold" }],
        "background": [{ "show": true, "color": { "solid": { "color": "#FFFFFF" } }, "transparency": 0 }],
        "border": [{ "show": false }],
        "dropShadow": [{ "show": false }],
        "padding": [{ "top": 8, "bottom": 8, "left": 10, "right": 10 }]
      }
    },
    "page": {
      "*": {
        "background": [{ "color": { "solid": { "color": "#F4F4F4" } }, "transparency": 0 }],
        "outspace":   [{ "color": { "solid": { "color": "#E9E9E9" } }, "transparency": 0 }]
      }
    },
    "pivotTable": {
      "*": {
        "columnHeaders": [{ "fontColor": { "solid": { "color": "#FFFFFF" } }, "backColor": { "solid": { "color": "#C21B17" } }, "fontFamily": "Segoe UI Semibold", "fontSize": 9, "alignment": "Right" }],
        "rowHeaders":    [{ "fontColor": { "solid": { "color": "#333333" } }, "backColor": { "solid": { "color": "#FFFFFF" } }, "fontFamily": "Segoe UI", "fontSize": 9 }],
        "values":        [{ "fontColor": { "solid": { "color": "#333333" } }, "backColor": { "solid": { "color": "#FFFFFF" } }, "backColorSecondary": { "solid": { "color": "#F7F7F7" } }, "fontFamily": "Segoe UI", "fontSize": 9 }],
        "grid":          [{ "gridHorizontal": true, "gridHorizontalColor": { "solid": { "color": "#E3E3E3" } }, "gridVertical": false, "outlineColor": { "solid": { "color": "#C21B17" } } }],
        "subTotals":     [{ "fontColor": { "solid": { "color": "#333333" } }, "backColor": { "solid": { "color": "#F2E3E2" } }, "fontFamily": "Segoe UI Semibold" }],
        "total":         [{ "fontColor": { "solid": { "color": "#333333" } }, "backColor": { "solid": { "color": "#F2E3E2" } }, "fontFamily": "Segoe UI Semibold" }]
      }
    },
    "tableEx": {
      "*": {
        "columnHeaders": [{ "fontColor": { "solid": { "color": "#FFFFFF" } }, "backColor": { "solid": { "color": "#C21B17" } }, "fontFamily": "Segoe UI Semibold", "fontSize": 9 }],
        "values":        [{ "fontColor": { "solid": { "color": "#333333" } }, "backColor": { "solid": { "color": "#FFFFFF" } }, "backColorSecondary": { "solid": { "color": "#F7F7F7" } }, "fontFamily": "Segoe UI", "fontSize": 9 }],
        "grid":          [{ "gridHorizontal": true, "gridHorizontalColor": { "solid": { "color": "#E3E3E3" } }, "gridVertical": false }],
        "total":         [{ "fontColor": { "solid": { "color": "#333333" } }, "backColor": { "solid": { "color": "#F2E3E2" } }, "fontFamily": "Segoe UI Semibold" }]
      }
    },
    "card": {
      "*": {
        "labels":         [{ "color": { "solid": { "color": "#C21B17" } }, "fontSize": 20, "fontFamily": "Segoe UI Semibold" }],
        "categoryLabels": [{ "show": true, "color": { "solid": { "color": "#6E6E6E" } }, "fontSize": 9, "fontFamily": "Segoe UI" }]
      }
    },
    "slicer": {
      "*": {
        "header": [{ "show": true, "fontColor": { "solid": { "color": "#C21B17" } }, "fontFamily": "Segoe UI Semibold", "textSize": 10 }],
        "items":  [{ "fontColor": { "solid": { "color": "#333333" } }, "fontFamily": "Segoe UI", "textSize": 9 }]
      }
    },
    "textbox": {
      "*": {
        "background": [{ "show": false }]
      }
    }
  }
}








______________________________


=====================================================================
 ДИЗАЙН ОТЧЁТА В СТИЛЕ GENERALI — пошагово
 Файл темы: Generali_CDD_Theme.json (прислан вместе с этой инструкцией)
=====================================================================

>> ПРИНЦИПЫ (чтобы выглядело профессионально)
>> 1 акцентный цвет (красный Generali) + серые. Красный = шапки, заголовки, акценты.
>>   Статусы (зелёный/жёлтый/красный) — только в точках-светофорах, больше нигде.
>> Одинаковая сетка на всех страницах: шапка, меню, контент — на тех же местах.
>> Без рамок и теней, белые карточки на светло-сером фоне, отступы 16 px.
>> Шрифт Segoe UI (корпоративные шрифты Generali в Power BI недоступны).
>> Красный в теме: #C21B17. Сверь с брендбуком на интранете (Brand portal) —
>> если там другой HEX, замени его в JSON поиском-заменой (Блокнот > Ctrl+H).

>> ПАЛИТРА
>>   Generali red   #C21B17   шапки таблиц, заголовки, главная серия на графиках
>>   Dark red       #7A0F0C   шапка страницы (полоса), акцентная линия
>>   Light red      #F2E3E2   строки Total / подсветка
>>   Text           #333333   основной текст
>>   Grey text      #6E6E6E   подписи, сноски
>>   Grid           #E3E3E3   линии таблиц
>>   Page           #F4F4F4   фон страницы
>>   Status: In line #2E7D32 | Watch #E0A100 | Action needed #A50E0E | Info #9E9E9E


#####################################################################
 ШАГ 1. ТЕМА (2 минуты, меняет все визуалы сразу)
#####################################################################
>> Visualizza > Temi (стрелка вниз) > Sfoglia temi > Generali_CDD_Theme.json > Apri.
>> Что сделает: шапки матриц/таблиц красные с белым текстом, полосатые строки,
>> серый фон страниц, без рамок, карточки с красной цифрой, слайсеры с красным заголовком.
>> Если какой-то визуал не поменялся — у него вручную задан формат:
>>   выдели визуал > Formato > (вверху) "Ripristina impostazioni predefinite".
>> Меры цвета (Summary Color, RAF Color, SAA Pct Color) продолжат работать.
>> В Summary Color поменяй красный на #A50E0E, чтобы статус не сливался с брендом:
>>   Summary Color = SWITCH ( [Summary Status], 1, "#2E7D32", 2, "#E0A100", 3, "#A50E0E", 0, "#9E9E9E" )


#####################################################################
 ШАГ 2. РАЗМЕР И СЕТКА СТРАНИЦЫ
#####################################################################
>> Каждая страница: Formato pagina (клик по пустому месту) > Impostazioni canvas >
>>   Tipo: 16:9 (1280 x 720).   Visualizza > Griglia + Blocca su griglia: вкл.
>> Сетка (координаты X, Y, ширина, высота — Formato > Generale > Proprietà):
>>
>>   ┌──────────────────────────────────────────────────────────────┐
>>   │ ШАПКА  0,0  1280x56  (тёмно-красная полоса)                  │
>>   ├──────────┬───────────────────────────────────────────────────┤
>>   │ МЕНЮ     │ ЛЕНТА ФИЛЬТРОВ  196,68  1068x44  (слайсеры в ряд)  │
>>   │ 0,56     ├──────────────────────────────┬────────────────────┤
>>   │ 180x664  │ ОСНОВНОЙ БЛОК                │ КОММЕНТАРИЙ        │
>>   │ (белое)  │ 196,124  760x580             │ 972,124  292x580   │
>>   │          │ (матрицы/графики)            │                    │
>>   └──────────┴──────────────────────────────┴────────────────────┘
>>   Сноска внизу основного блока: "Source: CDD v3.0 · EUR m · TEST DATA".


#####################################################################
 ШАГ 3. ШАПКА (делаешь 1 раз, потом копируешь)
#####################################################################
>> На странице Summary – Total:
>> 1) Inserisci > Forme > Rettangolo: X 0, Y 0, W 1280, H 56.
>>    Riempimento #7A0F0C, Bordo: off, Ombra: off.
>> 2) Inserisci > Casella di testo поверх полосы (X 196, Y 8, W 700, H 40):
>>    название страницы, Segoe UI Semibold 18, белый. Фон текстбокса: off.
>> 3) Справа в полосе — логотип: Inserisci > Immagine > официальный логотип с интранета
>>    (белая/негативная версия для тёмного фона), высота ~32 px, X ~1150, Y 12.
>>    Используй только официальный файл логотипа из брендбука.
>> 4) Подзаголовок с выбранной компанией (динамический). Мера:

Page Subtitle =
"Capital Deep Dive · " & SELECTEDVALUE ( FactData[Entity], "All companies" ) & " · EUR m"

>>    Scheda (nuova) или Scheda с мерой Page Subtitle под названием, белый текст 10,
>>    фон прозрачный, без подписи категории.
>> 5) Выдели прямоугольник + текст + логотип + подзаголовок > правый клик > Raggruppa >
>>    назови группу "Header" (Visualizza > Selezione).


#####################################################################
 ШАГ 4. БОКОВОЕ МЕНЮ (вместо полосы внизу)
#####################################################################
>> 1) Прямоугольник фона меню: X 0, Y 56, W 180, H 664, заливка #FFFFFF, без рамки.
>> 2) Существующий Navigatore pagine перемести внутрь: X 8, Y 72, W 164, H 640.
>>    Formato > Griglia layout > Orientamento: Verticale.
>>    Stile > Applica impostazioni a: Predefinito:
>>       Testo Segoe UI 9, #333333, allineamento a sinistra; Riempimento: off; Bordo: off.
>>    Stile > Stato: Al passaggio del mouse: Riempimento #F2E3E2.
>>    Stile > Stato: Selezionato: Riempimento #C21B17, Testo bianco, Segoe UI Semibold.
>>    Pagine: выключи скрытые и служебные страницы (DebugSAA и т.п.).
>> 3) Длинные имена вкладок укороти (меню берёт текст из имени вкладки), например
>>    "Sens on Net Result & Solvency" -> "Sensitivities". ВАЖНО: то же новое имя
>>    впиши в SummaryIndex[Page], иначе кнопка Summary не найдёт страницу.
>> 4) Сгруппируй фон + навигатор: группа "Menu".


#####################################################################
 ШАГ 5. СКОПИРОВАТЬ ШАПКУ И МЕНЮ НА ВСЕ СТРАНИЦЫ
#####################################################################
>> Выдели группы Header и Menu > Ctrl+C > на каждой странице Ctrl+V
>> (встают в те же координаты). Поменяй только текст названия страницы.
>> Старые заголовки страниц (крупный чёрный текст) и нижний навигатор удали.
>> Существующие визуалы сдвинь в сетку ШАГА 2 (Formato > Proprietà > X/Y/размер).


#####################################################################
 ШАГ 6. ВИЗУАЛЫ — ЕДИНЫЕ ПРАВИЛА
#####################################################################
>> Заголовки визуалов: sentence case, коротко ("Solvency ratio by year", не "SOL Ratio").
>> Матрицы/таблицы:
>>   - Totali: только там, где есть смысл; строка Total — фон #F2E3E2 (уже в теме).
>>   - Числа: EUR m без десятичных (#,##0), проценты 0.0%, п.п. +0.0 "pp".
>>   - Отрицательные: в теме нет — в каждой мере формат "#,##0;(#,##0)" при желании.
>>   - Ширина колонок: Formato > Intestazioni colonna > Ridimensionamento automatico: off,
>>     одинаковая ширина для годов.
>> Графики:
>>   - Главная серия #C21B17, вторая #5C5C5C (уже первые цвета темы).
>>   - Линии сетки: светло-серые, ось Y без названия, подписи данных — только на итогах.
>>   - Waterfall: Aumenta #5C5C5C, Riduci #C21B17, Totale #7A0F0C.
>> Карточки KPI: одна строка из 3-4 карточек одинакового размера (например 180x80),
>>   отступ 16 px между ними.
>> Слайсеры: стиль "Elenco a discesa" (Formato > Impostazioni filtro dei dati > Stile)
>>   в ленте фильтров — экономит место; заголовок Entity / Scale.


#####################################################################
 ШАГ 7. КОММЕНТАРИИ (самое "умное" — сделать заметным)
#####################################################################
>> Блок комментария (справа, 292x580):
>> 1) Прямоугольник-фон: белый, X 972, Y 124, 292x580.
>> 2) Тонкая красная полоса слева: прямоугольник X 972, Y 124, W 4, H 580, #C21B17.
>> 3) Заголовок "Key messages" (текстбокс, Segoe UI Semibold 11, #C21B17), X 984, Y 132.
>> 4) Сама карточка с мерой комментария ниже: Formato > Valore: Segoe UI 10, #333333,
>>    перенос текста вкл., выравнивание по левому краю; Etichetta categoria: off.
>> Одинаково на всех страницах (скопируй группу и поменяй меру).


#####################################################################
 ШАГ 8. SUMMARY-СТРАНИЦЫ (первое впечатление)
#####################################################################
>> Summary – Total, сверху вниз:
>> 1) Строка KPI (4 карточки): Solvency ratio (min over plan) | Total remittance |
>>    Free tangible capital (last year) | Areas needing action (кол-во красных).
>>    Мера для последней:

Summary Red Count =
COUNTROWS ( FILTER ( ADDCOLUMNS ( SummaryIndex, "@s", [Summary Status] ), [@s] = 3 ) ) + 0

>>    Цвет цифры: Formattazione condizionale > Colore > Regole: >0 -> #A50E0E, =0 -> #2E7D32.
>> 2) Summary Overview — широкий блок под KPI (как комментарий, с красной полосой).
>> 3) Таблица разделов: точка | Section | текст | Label; без сетки по вертикали,
>>    высота строк 28-32 px (Formato > Valori > Spaziatura), Label цветом статуса
>>    (Colore carattere > Valore campo > Summary Color).
>> 4) Кнопка "Open … →": Riempimento #C21B17, текст белый, скругление 4 px,
>>    стиль "Al passaggio del mouse" #7A0F0C.
>> Summary – Companies: тепловая карта — точки крупнее (UNICHAR 11044), заголовки
>>    колонок развернуть (Formato > Intestazioni colonna > A capo automatico: on).


#####################################################################
 ШАГ 9. ДЕТАЛИ, КОТОРЫЕ ДЕЛАЮТ "ДОРОГО"
#####################################################################
>> [ ] Одинаковые отступы: 16 px между всеми блоками (Visualizza > Griglia помогает).
>> [ ] Выравнивание: выдели несколько визуалов > Formato > Allinea > In alto / A sinistra.
>> [ ] Заголовки визуалов в одном стиле, без сокращений типа "SAA c Net SCR".
>> [ ] Подсказки (tooltip) по умолчанию: оставить; убрать "Conteggio di..." из визуалов.
>> [ ] Скрыть служебные страницы (DebugSAA, тестовые): правый клик по вкладке > Nascondi.
>> [ ] Скрыть технические таблицы/колонки в модели (AuditTool, DebugSAA, CMP_...):
>>     Visualizzazione modello > правый клик > Nascondi nella visualizzazione report.
>> [ ] Сноска на каждой странице: "Source: CDD v3.0 · amounts in EUR m · TEST DATA".
>> [ ] File > Opzioni > Report corrente: "Usa nuovi menu contestuali", снять
>>     "Mostra intestazione oggetti visivi" для публикации — меньше иконок над визуалами.

>> ПОРЯДОК: тема (5 мин) -> шапка + меню на Summary (30 мин) -> копия на все страницы
>> (30-40 мин) -> выравнивание визуалов по сетке (по 5-10 мин на страницу) -> Summary.
>> Пришли скрин Summary после ШАГОВ 1-4 — подправлю, прежде чем тиражировать.


              
