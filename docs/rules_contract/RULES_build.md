# Опись правил: `docx_build.py`, `docx_profile.py` (сборка DOCX и профиль оригинала)

Сводка агента-разведчика 20.10 (только чтение). Источники: GEO — геометрия блоков (geom/geo/frame/ys/xs), PROFILE —
статистика профиля `P[...]`, LAYOUT — раскладка листа `L[...]`, TAG — виды и флаги блоков, TEXT — регулярки,
TRANSLATION — перевод. Ни один тест не зовёт методы Builder по имени (косв. — через `Builder(...).build`);
`docx_profile.extract/check/centered/_flat/_list_firsts` тестами не вызываются.

| id | где | что решает | кат. | источники | откуда | тесты | пересекается |
|---|---|---|---|---|---|---|---|
| B1 | `_figure` L44 | абзац-рисунок: только картинки, одна выше 24 pt | STYLE | GEO, TAG | — | — | B37 (без порога), B93 |
| B2 | `is_label` L50 | строка-подпись: ≤80, вся жирная, «:» → Label UTS | STYLE | TAG, TEXT | 16.10 | — | B19, B37, B44, B93 |
| B3 | `is_code` L59 | код: >½ Consolas/Courier, не body_mono | STYLE | TAG | 3SZ19 | косв. | B32, B37, B93 |
| B4 | `merge_inline` L105 | склейка абзацев одной строки; подпись не клеится; пустые — прочь | TEXT | GEO, TAG, TEXT | JPDS 17.10 | — | B109 |
| B5 | `merge_inline` L136 | значок в «()»; пробел по промежутку на листе | TEXT | GEO, TEXT | ТЭ357 17.10 | косв. | B32, B33, B109 |
| B6 | `_span_starts` L193 | ячейки неполной строки по столбцам рамки (DP) | TABLE | GEO | ÜFK, ИН-1172, OAH 1.1 | 1 тест | B64 |
| B7 | `_autonum_ok`/`_head_dots`/build | автонумерация заголовков при ≥90 %; точка по большинству | LIST, INDENT | TEXT, TAG | UTS, ÜFK, Eskom | 2 теста | B21, B50, B105 |
| B8 | `_keep_tab` L269 | вернуть «⇥», потерянный переводом | TEXT | TEXT, TRANSLATION | UTS 16.10 | 1 тест | B36, B90 |
| B9 | `_col_align` L280 | выравнивание столбца голосованием (head/first/body) | TABLE | GEO | UTS 16.10, ÜFK 19.10 | 1 тест | B69 |
| B10 | `_verbatim`/`_restyle` L316 | начертание групп перевода без меток | TEXT | TAG, TRANSLATION | UTS 16.10 | 2 теста | B32 |
| B11 | `code_part` L355 | моноширинный: без венгерских букв — оригиналом; листинг построчно | TEXT | TEXT, TRANSLATION | UTS 15/19.10 | 2 теста | B32 |
| B12 | `Memory.add` L400 | метка пункта снимается из ключа памяти | LIST | TEXT | ÜFK I 16.10 | 1 тест | B55, B109 |
| B13 | `_styles` L484 | Normal: шрифт, интервалы, выключка, межстрочный | STYLE | PROFILE | 15.10 | косв. | B95, B97, B105 |
| B14 | `_styles` L491 | Heading 1–5: кегль, жирность, keep, page_break | STYLE, PAGE | PROFILE | Ф3 | косв. | B49, B99, B105 |
| B15 | `_styles` L504 | Caption/Figure UTS | CAPTION | PROFILE | Ф8 | — | B37, B103, B107 |
| B16 | `_styles` L510 | Table Text / Head / Small | TABLE, STYLE | PROFILE | — | — | B59, B66, B69 |
| B17 | `_styles` L516 | Code UTS | STYLE | PROFILE | — | косв. | B37, B39 |
| B18 | `_styles` L522 | List UTS / List Cell UTS | LIST, STYLE | PROFILE | ТЭ357 | — | B56, B105 |
| B19 | `_styles` L527 | Center UTS, Label UTS | STYLE | PROFILE | 16.10 | — | B2, B39, B44 |
| B20 | `_styles` L532 | зашитые значения UTS (Title Verdana 18, Front, Small, TOC, Header, Stamp) | STYLE | — | UTS | — | B71, B72, B82 |
| B21 | `_heading_numbering` L540 | нумерация Heading: табуляция и висячий = heads.tab | LIST, INDENT | PROFILE | UTS | косв. | B7, B100, B105 |
| B22 | `_list_ind` L571 | отступ по виду списка lists_kind | INDENT | PROFILE | ÜFK 17.10 | — | B23–B27, B52–B54, B102, B106 |
| B23 | `_list_ind` L577 | отступ по уровню lists, экстраполяция | INDENT | PROFILE | Ф7 | косв. | B22, B24, B101, B106 |
| B24 | `_list_ind` L575 | без профиля: 18·(2lv+1) / 36·(lv+1) | INDENT | — | — | — | B22, B23 |
| B25 | `_lvl_ind` L589 | римские: текст ≥ маркер+24 | INDENT | TAG | ÜFK | — | B22–B24, B106 |
| B26 | `_lvl_ind` L591 | знак-маркер: текст ≥ маркер+9 | INDENT | TAG | НГТУ 19.10 | — | B22–B24, B106 |
| B27 | `_list_num` L623 | w:ind уровня в numbering | INDENT, LIST | PROFILE, TAG | — | косв. | B53, B54, B106 |
| B28 | `_list_num` L610 | вид метки (Label/NUMFMT/маркеры, own) | LIST | TAG | ÜFK, 3SZ19, 3PR42 | — | B55 |
| B29 | `_list_num` L635 | startOverride — номер первого пункта | LIST | TAG | — | косв. | — |
| B30 | `_img_run` L647 | картинка не шире поля и не выше поля −48 | PAGE | LAYOUT, GEO | ИН-1065 18.10 | — | B63, B77 |
| B31 | `_own_size`/paragraph | прямой кегль при моде ≠ стиля на ≥1 pt | TEXT | TAG, LAYOUT | OAH 4.22 18.10 | косв. | B38, B71, B110 (М2) |
| B32 | `_put_runs` L692 | группы начертания, перевод целиком / по частям / с метками | TEXT | TAG, TRANSLATION | 3SZ19 | косв. | B3, B8, B10, B11 |
| B33 | `_put_runs` L685 | место значков (начало / at_tab / paren / конец) | TEXT | TAG, GEO | — | косв. | B5, B109 |
| B34 | `text_run` L755 | sub/sup, жирный, курсив, подчёркивание, Consolas, цвет | TEXT | TAG, TRANSLATION | 3PR42 | косв. | B50 |
| B35 | `_put_runs` L796 | outlineLvl пункту TOC без номера | TOC | TAG | — | косв. | B51, B72 |
| B36 | `_put_runs` L803 | табуляции по оригиналу; rtab — правая у поля | INDENT, TOC | GEO, LAYOUT, TAG | ÜFK II 18.10, 3SZ19 | — | B68, B90 |
| B37 | `paragraph` L823 | подпись → caption; картинки → Figure; код; Label | STYLE, CAPTION | TEXT, TAG | — | — | B1, B2, B3, B47, B93 |
| B38 | `paragraph` L835 | Title/Front/Small по кеглю относительно тела | STYLE | GEO, PROFILE | ИН-1065 16.10 | косв. | B71 (абсолютные 17/13/8), B31 |
| B39 | `paragraph` L846 | Center UTS, если `_centered` (только тело) | STYLE | GEO, LAYOUT | Ф13 | — | B48, B94, B109 |
| B40 | `paragraph` L848 | b.indent → Indent N | INDENT | TAG | docx_layout | — | B41–B43, B46 |
| B41 | `paragraph` L853 | висячий → Hanging d h | INDENT | GEO, LAYOUT | OAH | — | B40, B42, B46 |
| B42 | `paragraph` L859 | сдвиг тела → Indent d | INDENT | GEO, PROFILE, LAYOUT | OAH, МТ-14 16.10 | — | B40, B43, B84, B98 |
| B43 | `paragraph`/`_flush` L867 | у поля без красной строки → Body Flush UTS | INDENT | GEO, PROFILE | НГТУ 19.10 | 1 тест | B42, B84, B98 |
| B44 | `paragraph` L879 | keep_with_next: жирная строка-заголовок; «…:» под заголовком | PAGE | TAG, TEXT | OAH, ИН-1065, UTS, JPDS 18.10 | — | B14, B19 |
| B45 | `_footnote` L895 | сноска Word | TEXT | TAG, TEXT | — | 1 тест | B84 |
| B46 | `_indented` L989 | фабрика Indent/Hanging (наследуют красную строку Normal) | INDENT | — | — | — | B40–B42, B84 |
| B47 | `caption` L1002 | «Рисунок/Таблица» + SEQ, перевод названия | CAPTION | TEXT, TRANSLATION | — | — | B15, B37, B107 |
| B48 | `_centered` L1018 | центр абзаца (пункт с номером >40 — нет; титул — свой) | STYLE | GEO, LAYOUT, TEXT | ÜFK II 16.10 | — | B39, B89, B94 |
| B49 | `heading`/`_block` | b.center — заголовок по центру; b.break — с нового листа | STYLE, PAGE | TAG | НГТУ, docx_struct | — | B14, B75, B99 |
| B50 | `_heading` L1038 | текст заголовка: номер, цвет, b=False | TEXT, LIST | TAG, TEXT, TRANSLATION | ИН-1172, Eskom | 2 теста | B7, B34 |
| B51 | `_head_style` L1073 | Heading N NoTOC, если нет в оглавлении | TOC | TAG, TEXT | ÜFK | косв. | B83, B91 |
| B52 | `list_` L1097 | свой отступ own = geo − left (≠ профиля >3, маркер ≥−15) | INDENT | GEO, PROFILE | 3SZ19 19.10 | 2 теста | B22–B27, B53, B54 |
| B53 | `list_` L1130 | 1-й абзац пункта: numPr + прямые left/first_line | INDENT, LIST | GEO, PROFILE, TAG | НГТУ 19.10 | косв. | B27, B84, B110 |
| B54 | `list_` L1139 | следующие абзацы пункта: прямой left | INDENT | GEO, PROFILE | — | — | B53, B27 |
| B55 | `list_` L1118 | маркер LEAD снимается из текста | LIST | TEXT | — | 1 тест | B12, B28 |
| B56 | `list_`/`_block` | List UTS / List Cell UTS, уровень, картинка в пункте | LIST, CAPTION | TAG, TEXT | JPDS 17.10 | — | B18, B37 |
| B57 | `table`/`_borders`/`_cell_borders` | линии: nil / tcBorders / Table Grid; заливка | TABLE | GEO, TAG | UTS | косв. | B104, B108 |
| B58 | `_cell_margins`/`table` | поля ячеек; лимит по узкому столбцу | TABLE | GEO, PROFILE | ÜFK II 17.10 | косв. | B60, B63, B67, B104 |
| B59 | `table` L1189 | кегль мелких таблиц (Small, свой ≤8) | TABLE, STYLE | GEO | МТ-14 18.10 | — | B16, B60, B31 |
| B60 | `table` L1210 | сетка без линеек: свой кегль, поля 1 pt | TABLE | TAG | docx_layout._grids | — | B58, B59 |
| B61 | `table` L1227 | ширины по xs / левым краям | TABLE | GEO | ÜFK, Ф10 | — | B62, B63, B108 |
| B62 | `table` L1244 | безрамочная раскладка: последний столбец до поля, tblInd | TABLE, INDENT | GEO, LAYOUT | UTS стр. 93 | — | B61, B42 |
| B63 | `table`/`_fit_words` | столбец не уже картинки и самого длинного слова перевода | TABLE | GEO, TRANSLATION, PROFILE | JPDS, UTS, ÜFK 16–17.10 | косв. | B30, B58, B61 |
| B64 | `table` L1293 | объединения: span, ничей столбец, vcont | TABLE | GEO, TAG | — | — | B6 |
| B65 | `table` L1310 | шапка и tblHeader | TABLE | TAG | JPDS 18.10 | — | B16, B66, B108 |
| B66 | `table` L1316 | cantSplit; keepNext шапки/1-й строки | TABLE, PAGE | TEXT, TAG | ÜFK I 19.10 | — | B16, B65, B108 |
| B67 | `table` L1320 | высота строки по ys: пустая exact, row_lines atLeast | TABLE | GEO | ИН-1065, НГТУ 19.10 | — | B58, B59, B70 |
| B68 | `table` L1334 | _tab_base ячейки | INDENT | GEO | ÜFK II 18.10 | — | B36 |
| B69 | `table`/`_aligned` | выравнивание ячейки стилем (head/body/first, объединённая) | TABLE | GEO | ÜFK 16.10 | — | B9, B16, B18 |
| B70 | `table` L1356 | содержимое ячейки, пустая ячейка, абзац после таблицы | TABLE | TAG | UTS 16.10 | — | B56, B67, B75 |
| B71 | `_title_styles` L1437 | Title/Front/Small — мода абзацев ≥17 / 13–17 / ≤8 | STYLE | GEO, TAG | ÜFK 16.10 | косв. | B20, B38 |
| B72 | `_toc_styles` L1468 | заголовок оглавления/списка рисунков | TOC | TAG | — | косв. | B20, B35 |
| B73 | `_toc_styles` L1485 | «toc N»: отступ, висячий, табуляция, шаг | TOC, INDENT | TAG | UTS 19.10, OAH 1.1 | косв. | B36, B84 |
| B74 | `toc` L1510 | поле TOC / список рисунков | TOC | TAG | — | косв. | — |
| B75 | `_shrink_tail`/`_block` | пустой абзац перед разрывом — 1 pt | PAGE | TAG | ÜFK I 18.10 | — | B49, B70 |
| B76 | `_page`/`_num_start` | лист, поля, нумерация с числа оригинала | PAGE | LAYOUT, PROFILE | — | косв. | B95, B105, B86 |
| B77 | `_zone_table` L1597 | колонтитул-таблица | HEADER | LAYOUT | UTS, ÜFK | — | B79, B80, B82 |
| B78 | `_zone_text` L1661 | «N/NN», номер листа — поля PAGE/NUMPAGES | HEADER | TEXT, LAYOUT | «85. oldal» | — | — |
| B79 | `_zone_paras` L1697 | колонтитул без рамки — абзацами с табуляциями | HEADER | LAYOUT | — | — | B77, B81 |
| B80 | `_header` L1749 | вид шапки | HEADER | LAYOUT | UTS | — | B77, B79, B87 |
| B81 | `_header` L1784 | строки шапки, линия под шапкой | HEADER | LAYOUT | НГТУ | — | B79, B109 |
| B82 | `_footer`/build | штамп | HEADER, STYLE | LAYOUT | — | — | B20, B77 |
| B83 | build L1841 | сверка с оглавлением — при ≥60 % | TOC | TEXT | ИН-1172 16.10 | косв. | B51, B91 |
| B84 | build L1834 | красная строка Normal; прочим стилям — 0 | INDENT | PROFILE | МТ-14 | косв. | B42, B43, B46, B53, B98 |
| B85 | build L1874 | пустоты почти пустого листа и титула | PAGE | GEO, LAYOUT | OAH, N3a, Eskom 17.10 | — | — |
| B86 | build L1890 | титул и разделы, вводная часть | PAGE, HEADER | LAYOUT, GEO | ONS, ÜFK | косв. | B76 |
| B87 | build L1944 | логотипы обложки → шапка титула; фон — прочь | HEADER | GEO, LAYOUT | UTS, NSU | — | B80 |
| B88 | build L1955 | текст титула ниже поля → подвал титула | HEADER | GEO, LAYOUT | UTS, OAH 4.22 | косв. | B89 |
| B89 | build L1967 | титул: места абзацев и картинок по оригиналу | INDENT | GEO, LAYOUT | JPDS 16.10 | — | B48, B42, B110 |
| B90 | `_tab_groups` L2021 | общая табуляция пар «подпись ⇥ значение» по переводу | INDENT, TEXT | TRANSLATION, TAG | UTS, 3SZ19, ИН-1065 | — | B8, B36 |
| B91 | `_in_toc` L2093 | заголовок есть в оглавлении | TOC | TEXT | ÜFK | — | B51, B83 |
| B92 | `font_table` L76 | класс шрифта для замены | STYLE | TAG | — | 1 тест | — |
| B93 | profile `_flat` L24 | роль блока для профиля | STYLE | TAG, TEXT | 15.10 | — | B1–B3, B37 |
| B94 | profile `centered` L86 | по центру ±8, x0−left >24, ширина <0,85 | STYLE | GEO, LAYOUT | 3PR42 16.10 | — | B39, B48, B103, B109 |
| B95 | profile `extract` L116 | выключка both/left, P.page.right | STYLE, PAGE | GEO, LAYOUT | синтетика 16.10 | — | B13, B76, B105 |
| B96 | profile `extract` L117 | шрифт роли — мода | STYLE | TAG | Ф2 | — | B13–B17, B105 |
| B97 | profile `extract` L139 | межстрочный, интервалы — медианы | STYLE | GEO | Ф6 | — | B13, B14, B18, B19 |
| B98 | profile `extract` L178 | красная строка — медиана first_x0−x0 | INDENT | GEO | МТ-14, синтетика 16.10 | — | B42, B43, B84 |
| B99 | profile `extract` L186 | page_break уровня заголовков | PAGE | GEO | ÜFK 16.10 | — | B14, B105 |
| B100 | profile `extract` L199 | heads.tab | INDENT | GEO, TEXT | UTS | — | B21, B105 |
| B101 | profile `extract` L214 | lists[lv] — медианы (вкл. пункты в ячейках!) | INDENT | GEO | Ф7 | — | B23, B102, B106 |
| B102 | profile `extract` L228 | lists_kind[kind][lv] (без ячеек) | INDENT | GEO, TAG | ÜFK 17.10, ТЭ357 | — | B22, B101, B106 |
| B103 | profile `extract` L233 | центр рисунков/подписей, порядок подписи | CAPTION | GEO, TAG | Ф8 | — | B15, B107 |
| B104 | profile `extract` L242 | поля ячеек, заливка, цвет линий | TABLE | GEO | Ф9, Ф11 | — | B57, B58 |
| B105 | profile `check` L269 | Ф1–Ф6 | CHECK | PROFILE | 15.10 | — | B13, B14, B21, B76, B95–B100 |
| B106 | profile `check` L323 | Ф7: ind уровней против lists (own из B52 не видит) | CHECK | PROFILE | Ф7 | — | B22–B27, B101, B102 |
| B107 | profile `check` L338 | Ф8 | CHECK | PROFILE | Ф8 | — | B15, B47, B103 |
| B108 | profile `check` L364 | Ф9, Ф10 | CHECK | GEO | — | — | B57, B61, B65, B66 |
| B109 | profile `check` L385 | Ф12–Ф15 | CHECK | GEO, TEXT | — | 1 тест | B5, B12, B33, B39, B81, B94 |
| B110 | profile `check` L409 | М2: прямой кегль/отступ (кроме List UTS) | CHECK | — | метрика 2 | — | B31, B53, B89 |

## Пересечения (по сводке агента)

1. **Отступы пунктов списка — 17 правил.** Профиль B101 (с ячейками) и B102 (без ячеек) — расходятся. Выбор: B22 →
   B23 → B24 (`_list_ind`), поправки B25/B26 (`_lvl_ind`). Значение пишется дважды: в numbering (B27) и прямо в абзац
   (B53, B54). B52 подменяет отступ на own из geo — только в абзаце, в numbering остаётся профильный. B84 обнуляет
   красную строку List UTS — потому B53 ставит прямые отступы. Уровень — B56 (b.level), B12/B55. Проверки: B106 (Ф7)
   видит только numbering (own не видит, B25/B26 дают ложный Ф7); B110 (М2) освобождает только «List UTS» — прямые
   отступы пунктов в ячейках считаются нарушением. Рядом: отступ номера заголовков B100 → B21 → B7 → B105 (Ф4).
2. **Отступ абзаца и красная строка** — B98 → B84; B40, B41, B42, B43, B46; титул B89; таблица B62; оглавление B73;
   нули B45, B79; B7; проверка B110.
3. **Высота строк таблицы** — B67, B58, B60, B59, B16, B70, B66, B75.
4. **Выравнивание ячеек** — B9, B69, B16, B18, B39 (в ячейке не применяется); колонтитулы B77, B79, B80.
5. **Центрирование абзацев** — B94 и B48 (два правила) → B39; B49; B15/B103; B88/B89; проверка B109.
6. **Кегль и «крупные» стили** — B38 (от кегля тела) и B71 (абсолютно 17/13/8) расходятся при теле 14 pt; B31 ставит
   прямой кегль, B110 (М2) считает это нарушением; таблицы — B59/B60.
7. **Удержание и разрывы** — B14, B49, B44, B15, B19, B16/B65/B66, B75, B85.
8. **Ширины столбцов** — B61, B62, B63, B58, B6/B64; проверка B108.
9. **Правое поле** — B95 → B76 → B105.
10. **Признак «рисунок»** — B1 (выше 24 pt) и B37 (без порога) решают по-разному.
