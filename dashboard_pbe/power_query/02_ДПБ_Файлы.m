// ДПБ_Файлы — ДЗ покупателей по ПБЕ из папки «ДЗ и КЗ».
// Исправления по сравнению со старым запросом ДЗКЗ_ПБЕ:
//   1) строки с пустым итоговым сальдо НЕ выбрасываются — это контрагенты, рассчитавшиеся за месяц полностью,
//      их обороты должны входить в оборот (иначе оборот занижен в разы, а DSO завышен);
//   2) период = конец месяца даты из имени файла; берутся только файлы на 1-е или последнее число месяца
//      («01.03.2024» = конец марта), промежуточные срезы (29.05.2024, 09.04.2025, 07.05.2025) пропускаются;
//      если на месяц два файла — берётся файл с более поздней датой в имени.
let
    ПапкаИсточник = "C:\Users\v.lis\Desktop\ФинЦикл\Расчёт и анализ ФЦ\ДЗ и КЗ",
    Источник = Folder.Files(ПапкаИсточник),
    ФильтрФайлов = Table.SelectRows(Источник, each Text.StartsWith([Name], "ДЗ и КЗ")
        and (Text.EndsWith(Text.Lower([Name]), ".xlsx") or Text.EndsWith(Text.Lower([Name]), ".xls"))),

    ДатаИзИмени = (name as text) as nullable date =>
        let
            БезРасширения = Text.BeforeDelimiter(name, ".", {0, RelativePosition.FromEnd}),
            Токены = Text.Split(БезРасширения, " "),
            Кандидаты = List.Select(Токены, (т) =>
                let ч = Text.Split(т, ".") in
                List.Count(ч) = 3 and List.AllTrue(List.Transform(ч, (x) => (try Number.FromText(x) otherwise null) <> null))),
            Результат = if List.Count(Кандидаты) = 0 then null
                else let Ч = Text.Split(List.Last(Кандидаты), ".") in
                    try #date(Number.FromText(Ч{2}), Number.FromText(Ч{1}), Number.FromText(Ч{0})) otherwise null
        in
            Результат,
    СДатой = Table.AddColumn(ФильтрФайлов, "ДатаФайла", each ДатаИзИмени([Name]), type nullable date),
    Годные = Table.SelectRows(СДатой, each [ДатаФайла] <> null
        and (Date.Day([ДатаФайла]) = 1 or [ДатаФайла] = Date.EndOfMonth([ДатаФайла]))),
    СПериодом = Table.AddColumn(Годные, "Период", each Date.EndOfMonth([ДатаФайла]), type date),
    Лучшие = Table.Group(СПериодом, {"Период"}, {{"МаксДата", each List.Max([ДатаФайла]), type date}}),
    Соед = Table.NestedJoin(СПериодом, {"Период"}, Лучшие, {"Период"}, "Л", JoinKind.Inner),
    Раскр = Table.ExpandTableColumn(Соед, "Л", {"МаксДата"}),
    ОдинФайл = Table.SelectRows(Раскр, each [ДатаФайла] = [МаксДата]),

    Норм = (t as any) as text => Text.Replace(Text.Lower(if t = null then "" else Text.From(t)), "ё", "е"),
    ВыбратьЛист = (bookContent as binary) as table =>
        let
            Книга = Excel.Workbook(bookContent, true),
            Проверить = (data as table) as nullable table =>
                let
                    ЕстьСразу = List.AnyTrue(List.Transform(Table.ColumnNames(data), each Text.Contains(Норм(_), "счет"))),
                    ПослеПромо = if ЕстьСразу then null else (try Table.PromoteHeaders(data, [PromoteAllScalars = true]) otherwise null),
                    ЕстьПослеПромо = if ПослеПромо = null then false
                        else List.AnyTrue(List.Transform(Table.ColumnNames(ПослеПромо), each Text.Contains(Норм(_), "счет")))
                in
                    if ЕстьСразу then data else if ЕстьПослеПромо then ПослеПромо else null,
            СПроверкой = Table.AddColumn(Книга, "Checked", each try Проверить([Data]) otherwise null),
            Подходящие = Table.SelectRows(СПроверкой, each [Checked] <> null),
            Результат = if Table.RowCount(Подходящие) = 0 then error "Не найден лист со столбцом 'Счет'"
                else Подходящие{0}[Checked]
        in
            Результат,
    НайтиСтолбец = (t as table, содержит as list, неСодержит as list) as text =>
        let
            Совпадения = List.Select(Table.ColumnNames(t), (имя) =>
                let низ = Норм(имя) in
                List.AllTrue(List.Transform(содержит, (с) => Text.Contains(низ, с)))
                and List.AllTrue(List.Transform(неСодержит, (с) => not Text.Contains(низ, с)))),
            Результат = if List.Count(Совпадения) > 0 then Совпадения{0}
                else error ("Столбец не найден: " & Text.Combine(содержит, "+") & ". Реальные столбцы: " & Text.Combine(Table.ColumnNames(t), " | "))
        in
            Результат,
    НайтиПредпочтительно = (t as table, содержит as list) as text =>
        let без = try НайтиСтолбец(t, содержит, {"эквивал"}) otherwise null
        in if без <> null then без else НайтиСтолбец(t, содержит, {}),
    ЧитатьФайл = (bin as binary) as table =>
        let
            Лист = ВыбратьЛист(bin),
            СтСчет = НайтиСтолбец(Лист, {"счет"}, {"открытия"}),
            СтСальдо = НайтиПредпочтительно(Лист, {"итоговое", "сальдо"}),
            СтДт = НайтиПредпочтительно(Лист, {"оборот", "дебет"}),
            СтКт = НайтиПредпочтительно(Лист, {"оборот", "кредит"}),
            СтПБЕ = try НайтиСтолбец(Лист, {"пбе"}, {}) otherwise null,
            Выбрано = Table.SelectColumns(Лист, List.RemoveNulls({СтСчет, СтСальдо, СтДт, СтКт, СтПБЕ})),
            Переим = Table.RenameColumns(Выбрано, List.RemoveNulls({{СтСчет, "Счет"}, {СтСальдо, "Сальдо"},
                {СтДт, "Дт"}, {СтКт, "Кт"}, if СтПБЕ = null then null else {СтПБЕ, "ПБЕ"}})),
            СПБЕ = if СтПБЕ = null then Table.AddColumn(Переим, "ПБЕ", each null) else Переим
        in
            СПБЕ,

    СДанными = Table.AddColumn(ОдинФайл, "Данные", each ЧитатьФайл([Content])),
    Нужное = Table.SelectColumns(СДанными, {"Период", "Данные"}),
    Раскрыто = Table.ExpandTableColumn(Нужное, "Данные", {"Счет", "Сальдо", "Дт", "Кт", "ПБЕ"}),
    Типы = Table.TransformColumnTypes(Раскрыто, {{"Счет", type text}, {"Сальдо", type number}, {"Дт", type number}, {"Кт", type number}}),
    БезОшибок = Table.RemoveRowsWithErrors(Типы, {"Сальдо", "Дт", "Кт"}),
    Нули = Table.ReplaceValue(БезОшибок, null, 0, Replacer.ReplaceValue, {"Сальдо", "Дт", "Кт"}),
    СчетаДЗ = {"1210", "1211", "1211(КОНС)", "1212", "1213", "1215", "1260"},
    ТолькоДЗ = Table.SelectRows(Нули, each [Счет] <> null and List.Contains(СчетаДЗ, Text.Trim([Счет]))),
    ПБЕЧист = Table.TransformColumns(ТолькоДЗ, {{"ПБЕ", each
        if _ = null or Text.Trim(Text.From(_)) = "" then "(без метки ПБЕ)" else Text.Trim(Text.From(_)), type text}}),
    Итог = Table.Group(ПБЕЧист, {"Период", "ПБЕ"}, {
        {"ДЗТг", each List.Sum(List.Select([Сальдо], (x) => x > 0)), type number},
        {"АвансыТг", each - List.Sum(List.Select([Сальдо], (x) => x < 0)), type number},
        {"ОборотДтТг", each List.Sum([Дт]), type number},
        {"ОборотКтТг", each List.Sum([Кт]), type number}}),
    ИтогНули = Table.ReplaceValue(Итог, null, 0, Replacer.ReplaceValue, {"ДЗТг", "АвансыТг", "ОборотДтТг", "ОборотКтТг"})
in
    Table.Buffer(ИтогНули)
