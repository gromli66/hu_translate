# Опись правил: `src/docx_struct.py` (PDF с разметкой Word)

Сводка агента-разведчика 20.10 (только чтение кода). Источники: GEO — геометрия листа, TAG — теги структуры PDF,
TEXT — регулярки по тексту, STYLE — начертание. PORTAL и PROFILE в модуле не используются.

| id | где | что решает | категория | источники | откуда | тесты | пересекается |
|---|---|---|---|---|---|---|---|
| S1 | `_upright` L34 | знак повёрнут — по матрице | HEADER | GEO | JPDS | — | S42, S86 |
| S2 | `_underline` L40 | подчёркивание: линия ≤1,2 pt под базовой внутри рамки знака | STYLE | GEO | OAH 4.22, ÜFK | test_underline_from_thin_rule_under_glyphs | — |
| S3 | `_fmt` L62 + BOLD_W/ITAL_W | полужирный/курсив по имени шрифта (вкл. «félkövér»), курсив по наклону | STYLE | STYLE, GEO | ÜFK 16.10 | test_slanted_tahoma_is_italic… | S93, S97, S100, S90 |
| S4 | `_runs` L100 | Math Alphanumeric → NFKC | TEXT | TEXT | ИН-1172 17.10 | test_math_letters_become_plain_text | — |
| S5 | `_runs` L104 | мелкий знак выше/ниже → над-/подстрочный Юникод | STYLE | GEO, STYLE | ИН-1065, ÜFK II 17.10 | test_raised_and_lowered_small_chars… | docx_build (OCR sub/sup) |
| S6 | `_runs` L113 | переход строки внутри элемента → пробел | JOIN | GEO | ÜFK II 17.10 | косв. | S24, S12, S99 |
| S7 | `_runs` L120 | промежуток >1,8 кегля / знак на висячем отступе → «\t» | INDENT | GEO | — | test_justified_line_with_real_tab_keeps_tab, test_hanging_tab_and_stretch_after_it | S8, S9, S10, S70 |
| S8 | `_runs` L133 | ≥3 табуляций → выключка, «\t» → пробелы | INDENT | TEXT | — | косв. | S9 |
| S9 | `_justified_gaps` L152 | широкий промежуток — растяжка (медиана после висячей табуляции) | INDENT | GEO | OAH 16.10, 19.10 | 3 теста | S7, S8, S10 |
| S10 | `_hanging_tab` L186 | знак на висячем отступе после >0,5 кегля → табуляция | INDENT | GEO | OAH 18.10 | 2 теста | S7, S9, S42 |
| S11 | `_tail_pairs` L211 | хвост строки таблицы на новом листе | TABLE | GEO, TAG, TEXT | синтетика 16.10, JPDS | косв. | S18, S24, S37 |
| S12 | `_cut_rows` L234 | хвост — в ячейки строки выше | JOIN | TEXT, TAG | синтетика 16.10 | test_row_cut_by_page_tail… | S24, S37 |
| S13 | `insert_imgs` L269 | значок в строке — перед прогоном правее | FIGURE | GEO | UTS 16.10 | test_icon_in_line… | S31, S70 |
| S14 | `fit_rows` L286 | лишние ячейки одного столбца рамки — одна | TABLE | GEO | JPDS т.15 19.10 | — | S45 |
| S15 | `sub_rows` L326 | метки с пустыми абзацами → подстроки, vcont | TABLE | GEO, TEXT | ÜFK 19.10 | test_padded_labels_become_sub_rows | S81 |
| S16 | `join_tables` L399 | пометка «Folytatás…» между частями таблицы — прочь | TABLE | GEO, TEXT, TAG | ÜFK 16.10 | test_spans_autonum… | S96 |
| S17 | `join_tables` L390 | повтор шапки (TH или дословный) — снять | TABLE | TAG, TEXT | синтетика, JPDS | 2 теста | docx_build tblHeader |
| S18 | `join_tables` L416 | продолжение таблицы на след. листе — одна таблица, общие ys | TABLE | TAG, GEO | UTS 16.10, ИН-1065 | 3 теста | S11, S43, S45 |
| S19 | `join_tables`+`adopt` | значки из повтора шапки — в поток (strays) | FIGURE | GEO | JPDS т.15 | 1 тест | S21, S22 |
| S20 | `attach_notes` L445 | сноска — к абзацу со знаком ссылки | ROLE | TAG, TEXT, STYLE | ТЗ ВШЭ 16.10 | 1 тест | S59, S26 |
| S21 | `adopt` L551 | абзац в полосе строки таблицы — в ячейку | TABLE | GEO | ТЭ357 т.16, JPDS т.65 | 1 тест | S23, S77 |
| S22 | `adopt` L633 | значок в чужой строке — в свою | FIGURE | GEO | JPDS т.65 | 1 тест | S19 |
| S23 | `evict` L643 | блок ячейки ниже следующего за таблицей — в поток | TABLE | GEO | JPDS 16–17.10 | 2 теста | S21, S68, S72 |
| S24 | `join_pages` L709 | абзац через лист — один (строчная / красная строка / строка во всю ширину) | JOIN | GEO, TEXT | синтетика, OAH 1.1 19.10 | 2 теста | S12, S37, merge_inline |
| S25 | `join_pages` L717 | продолжение с меткой пункта — не склеивать | JOIN | TEXT | OAH N3a 19.10 | 1 тест | S33, S32 (свои регулярки) |
| S26 | `join_notes` L732 | сноски склеенных абзацев | JOIN | TEXT | ревью 17.10 | 1 тест | S20 |
| S27 | `reflow` L744 | блок далеко от своего листа — к своему листу | PAGE | GEO | ТЭ357 | 1 тест | S72, S38 |
| S28 | `join_tocs` L788 | оглавления подряд — одно | TOC | TAG | ÜFK 16.10 | 2 теста | S48–S51 |
| S29 | `join_tocs` L797 | короткий абзац перед TOC — его заголовок | TOC | TEXT, GEO, STYLE | OAH 4.22 18.10 | — | S48 |
| S30 | `_centered_geom` L813 | по центру: ±30 от 297,6 (жёстко A4) | STYLE | GEO | — | — | S48, S92, S97/S100 |
| S31 | `attach_icons` L832 | картинка-артефакт ≤20 pt в строке — в абзац | FIGURE | GEO, TAG | UTS 16.10 | 1 тест | S13, S70, S84/S96 |
| S32 | `text_labels` L855 | вид списка по меткам в тексте | ROLE | TAG, TEXT | ÜFK 16.10 | 2 теста | S33, S39, LEAD |
| S33 | `para_lists` L907 | абзацы с метками подряд → список | ROLE | TEXT, GEO | ÜFK I/II 16.10 | 1 тест | S32, S25 |
| S34 | `para_lists` L929 | «f.)» после «a.)–e.)» — в тот же список | ROLE | TEXT, GEO | ÜFK 16.10 | 1 тест | S36 |
| S35 | `_strip_label`+`relist` | номер пункта текстом — из прогонов прочь | TEXT | TEXT, TAG | UTS | косв. | LEAD, S20 |
| S36 | `_nest` L995 | новый список при смене вида/номера, start | ROLE | TEXT, TAG | — | 1 тест | S34, S69 |
| S37 | `_continued` L1019 | пункт через лист со строчной — продолжение | JOIN | TEXT, GEO, TAG | OAH 16.10 | 1 тест | S24, S12 |
| S38 | `relist` L1055 | пункты списков — по (лист, верх) | ROLE | GEO, TAG | UTS 44–45 | 1 тест | S68, S72 |
| S39 | `relist` L1068 | вид пункта по маркеру/номеру в тексте | ROLE | TEXT, TAG | синтетика, OAH | 2 теста | S32, S56, LEAD |
| S40 | `relist` L1084 | своя метка без Lbl → Label | ROLE | TEXT, TAG | OAH T0.7, N3a 18.10 | — | S33, docx_build Label |
| S41 | `load` L1108 | артефакт / содержимое по MCID | HEADER | TAG | — | косв. | S82, S31 |
| S42 | `geom_of` L1138 | x0/x1/first_x0/rest_x0/lines/last_x1/rotated/size | INDENT | GEO | OAH 16.10 | косв. | S63, paragraph, profile |
| S43 | `blank_geom` L1159 | таблица без текста — место по линиям | TABLE | GEO | ИН-1065 19.10 | — | S62, S18 |
| S44 | `borders` L1175 | рамка таблицы есть, если ≥2 линий | TABLE | GEO | — | — | S81, S45 |
| S45 | `frame` L1192 | столбцы xs (вертикаль ≥½ высоты), низ по цепочке, ys | TABLE | GEO | JPDS т.15, ИН-1065 19.10 | косв. | S14, S43, S18 |
| S46 | `frame` L1239 | цвет линий, заливка шапки TH | TABLE | GEO, TAG, STYLE | — | — | S80, S78 |
| S47 | `frame` L1249 | поля ячеек | TABLE | GEO | N3a 17–18.10 | — | profile |
| S48 | `toc_info` L1288 | заголовок оглавления, центр ±30 от W/2 | TOC | TEXT, GEO | — | — | S29, S30 |
| S49 | `toc_info` L1290 | уровень пункта TOC: глубина номера / кластеры отступа | TOC | TEXT, GEO, STYLE | ТЭ357 16.10 | — | S65, S66 |
| S50 | `toc_info` L1312 | список рисунков, если >½ «N. ábra» | TOC | TEXT | — | — | S67 |
| S51 | `toc_info` L1313 | extra/keys/entries | TOC | TEXT | UTS, ÜFK 16.10 | косв. | S28, S66 |
| S52 | `fills`/`thin` L1342 | заливка (не белая <0,97), тонкая линия | TABLE | GEO, STYLE | OAH 4.22 18.10 | — | S77/S78 (0,95), S45, S97, S98 |
| S53 | `block` L1377 | H1–H6 → заголовок уровня тега | ROLE | TAG | — | косв. | S64–S66 |
| S54 | `block` L1379 | P/Caption/BlockQuote → абзац | ROLE | TAG | — | косв. | S67 |
| S55 | `block` L1381 | Figure → figure=True | FIGURE | TAG | — | — | S72–S76 |
| S56 | `block` L1383 | L → список (ListNumbering, Lbl, LBody) | ROLE | TAG | — | — | S32, S33, S39 |
| S57 | `block` L1395 | Table: TR/TD/TH, th, место, edges/fills/frame | TABLE | TAG, GEO | — | — | S17, S46 |
| S58 | `block` L1422 | TOC — место оглавления | TOC | TAG | ÜFK | — | S28 |
| S59 | `block` L1427 | Note → блок сноски | ROLE | TAG | — | косв. | S20 |
| S60 | `block` L1430 | Link/Span/… вне абзаца → inline | JOIN | TAG | — | — | merge_inline, S24 |
| S61 | `block`/`container` L1434 | свои MCID обёрток → абзац | ROLE | TAG | UTS 16.10 | — | S31 |
| S62 | `load` L1455 | верх/низ содержимого листа; пустой — по линиям 12–92 % | PAGE | GEO, TAG | ИН-1065 18.10 | — | S43, S79 |
| S63 | `list_geo` L1475 | место метки и текста списка — медианы | INDENT | GEO | ÜFK 17.10 | 1 тест | S42, S83, profile, `_list_ind`, docx_layout (копия) |
| S64 | `real_headings` L1499 | заголовок по тегу → абзац, если длинный/не жирный | ROLE | TAG, TEXT, GEO, STYLE | ИН-1172 18.10 | 1 тест | S53 |
| S65 | `real_headings` L1510 | уровень — по глубине номера | ROLE | TEXT, TAG | ИН-1172 | 1 тест | S49, docx_build |
| S66 | `toc_headings` L1518 | заголовок-картинка — текст пункта TOC | TOC | TEXT, TAG | JPDS | 1 тест | S51, S75 |
| S67 | `CAPTION_ANY`/`FIG_CAPTION` | подпись по тексту (две регулярки) | ROLE | TEXT | — | — | S54, S69, S71–S73, docx_build |
| S68 | `list_tails` L1549 | хвост пункта ниже следующего абзаца — за него | FIGURE | GEO | JPDS л.42 17.10 | 1 тест | S71, S72, S23 |
| S69 | `unlist_captions` L1591 | пункт-подпись — не пункт | FIGURE | TEXT | JPDS 17.10 | — | S36, docx_build |
| S70 | `inline_pictures` L1620 | картинка в высоте абзаца — в него на табуляцию | FIGURE | GEO, TEXT | ИН-1065 17.10, UTS 18.10 | 2 теста | S13, S31, S7 |
| S71 | `place_captions` L1649 | «Рисунок N» до картинки, а на листе под ней → за неё | FIGURE | GEO, TEXT | JPDS 17.10 | — | S72, S68 |
| S72 | `figures_in_place` L1691 | рисунок — по месту на листе | FIGURE | GEO, TEXT | JPDS 17.10, ТЭ357 | 1 тест | S71, S68, S27, S23 |
| S73 | `_compose_figures` L1779 | рост области рисунка (векторы, указатели, короткие надписи) | FIGURE | GEO, TEXT | ТЭ357 17/19.10 | 1 тест | S77, S72 |
| S74 | `_compose_figures` L1843 | картинка с короткими надписями / мелкая с векторами | FIGURE | GEO, TEXT | JPDS л.96, ТЭ357 19.10 | — | S72 |
| S75 | `_compose_figures` L1879 | картинка вне разметки (XObject) | FIGURE | GEO | JPDS л.96 | — | S77, S31 |
| S76 | `_compose_figures` L1900 | рисунок из векторов | FIGURE | GEO | ÜFK I л.48, ONS | 1 тест | S89 |
| S77 | `_cell_pictures` L1958 | картинка ячейки вне разметки | TABLE | GEO, TEXT | JPDS 17.10 | 1 тест | S21, S73, S75 |
| S78 | `_cell_pictures` L1995 | пустая ячейка с заливкой — картинкой | TABLE | GEO, STYLE | JPDS 17.10 | 1 тест | S80, S46 |
| S79 | `mark_breaks` L2067 | разрыв листа (>72 pt выше обычного низа) | PAGE | GEO | ÜFK, Eskom, UTS, ИН-1065, OAH | — (зовёт и docx_layout) | S62, S83 |
| S80 | `cell_fills` L2108 | заливка ячейки | TABLE | GEO, STYLE | OAH 4.22 18.10 | — | S46, S78 |
| S81 | `cell_edges` L2123 | линии ячейки | TABLE | GEO | ÜFK 16.10 | 1 тест | S44, S15, S97 |
| S82 | `layout` L2182 | без разметки: тело / колонтитул по рамкам | HEADER | GEO | Eskom | — | S41 |
| S83 | `layout` L2204 | поля листа | PAGE | GEO | OAH, МТ-14, ТЭ357 16.10 | — | S63, S79, paragraph |
| S84 | `layout` L2241 | подвал/шапка: сетка по линиям / абзацы / сетка по разрывам | HEADER | GEO | OAH, UTS, НГТУ, ÜFK | косв. | S92, S97, S98 |
| S85 | `layout` L2287 | почти пустой лист — пустоты по оригиналу | PAGE | GEO, TEXT | — | — | — |
| S86 | `layout.zone` L2294 | 1-й лист — титул | PAGE | GEO, TEXT, STYLE | HSE, RS, ONS, UTS, JPDS 16.10 | — | S1, S94, docx_layout.layout |
| S87 | `layout` L2339 | колонтитулы 1-го листа, вводная часть — свой раздел | HEADER | GEO, TEXT | NSU, ÜFK, OAH 16.10 | — | S86 |
| S88 | `layout` L2237 | номер листа → поле PAGE, смещение нумерации | HEADER | GEO, TEXT | синтетика, ÜFK 16.10 | — | S95 |
| S89 | `_vector_logos` L2362 | логотип шапки из векторов | HEADER | GEO | ONS 17.10 | — | S76 |
| S90 | `font_classes` L2391 | класс шрифта по флагам | STYLE | STYLE | OAH 16.10 | — | S3 |
| S91 | `_head_rects` L2410 | линии зоны шапки | HEADER | GEO | синтетика 16.10 | — | S96, S83 |
| S92 | `_zone_paras` L2426 | колонтитул без рамок — абзацами (разрыв >18 pt, центр ±10) | HEADER | GEO | OAH 16.10 | 1 тест | S97, S98, S30 |
| S93 | `_zone_paras` L2459 | строка колонтитула полужирная — все знаки bold | STYLE | STYLE | Eskom 17.10 | 1 тест | S3, S97, S100 |
| S94 | `_typical_page` L2483 | типичный лист для колонтитулов | HEADER | TEXT, GEO | ÜFK, синтетика 16.10 | — | S86 |
| S95 | `_typical_page` L2497 | номер листа — слово с постоянным смещением | HEADER | TEXT, GEO | OAH, ÜFK | — | S88 |
| S96 | `_foot_top` L2518 | верх подвала | HEADER | GEO | синтетика, ÜFK I, Eskom | — | S91, S16, S31, S62 |
| S97 | `_rule_grid` L2562 | колонтитул-таблица по линиям | HEADER | GEO, STYLE | Печора, UTS, ÜFK | — | S98, S100, S81 |
| S98 | `_zone_grid` L2636 | столбцы колонтитула (кластер 8 pt, разрыв >18 pt) | HEADER | GEO | UTS 16.10 | — | S92, S97 |
| S99 | `_zone_grid` L2674 | перенос внутри ячеек колонтитула | JOIN | GEO | синтетика 16.10 | — | S6, S24 |
| S100 | `_zone_grid` L2666 | полужирная строка штампа — только «Bold» в имени | STYLE | STYLE, GEO | ÜFK | — | S3, S93, S97 |

## Пересечения (по сводке агента)

- **Табуляция или растяжка:** S7–S10; S70 опирается на позиции табуляций; в `docx_build` — свой `_keep_tab`.
- **Склейка разрезанного листом/строкой:** строки S6; строки таблиц S11, S12; таблицы S18 (+S16, S17); абзацы S24–S26;
  пункты S37; оглавление S28; штамп S99; строчные элементы S60 + `merge_inline`.
- **Заголовок/абзац и уровень:** S53, S64, S65, S49 (та же идея — своя реализация), S66, S51, S29.
- **Вид списка и метки:** S56, S32–S34, S39, S40, S36, S35, S25, S69 + `docx_build.LEAD`. Регулярки меток расходятся:
  LABEL, PLABEL, DLABEL, OWN_LABEL, DASH, регулярка S25, регулярка маркера S39.
- **Отступы списков и абзацев:** S42, S63, S83, S10 + `docx_profile`, `docx_build._list_ind`/`paragraph`, копия в
  `docx_layout`.
- **Перестановка блоков по месту:** S19, S21–S23, S27, S38, S68, S71, S72 — у каждого свои допуски (2, 30, 50 pt);
  итог зависит от порядка цепочки в `load`.
- **Значки в строке:** S13, S31, S70, S19, S22.
- **Рисунок:** S55, S72–S78. **Подписи:** S67 (две регулярки), S54 (тег Caption не используется), S69, S71–S73.
- **Таблицы:** S14 и S45 — одна причина, два правила; S21 и S77 — две реализации «полоса строки × столбцы»; S15
  перезаписывает edges S81; S44 (флаг рамки) и S81 (рамка по ячейкам).
- **Заливки:** S46, S80, S78, S52; порог «белого» — 0,97 (S52) и 0,95 (S73, S77, S78).
- **Граница зоны колонтитулов:** 12/88 % (S31), 12/92 % (S43, S62), 0,6H/0,7H (S96), 0,8H (S84), 20-й процентиль
  (S83), S91 — разные пороги одной границы.
- **Центровка:** ±30 от 297,6 — жёстко A4 (S30); ±30 от W/2 (S48); ±10 (S92); max(6, 5 %) (S97, S100).
- **Полужирный:** S3 (BOLD_W с «félkövér»), S93, S97, S100 (только «Bold» — «félkövér» не распознаётся), S49, S64.
- **Разрывы:** S79 (и из `docx_layout`), S62, S83, S86, S85. **Номер листа:** S88, S95.
- **«Короткий текст»:** ≤40 знаков (S29, S72–S74, S77), ≤20 (S15), <60 (S16).
