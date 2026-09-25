// Имя запроса: fnДП_Остатки   (это ФУНКЦИЯ — вставляется в пустой запрос целиком)
// Загрузка: «Только создать подключение»
//
// Читает помесячные файлы остатков из папки «Запасы» и возвращает таблицу
//   Период | Представительство | СуммаТг
// Отличия от старого запроса «Запасы_Представительства»:
//   1) ОДНА дата конца месяца = ОДИН файл (самый свежий по дате изменения).
//      Если в папку случайно положат второй файл с тем же месяцем — суммы НЕ удвоятся.
//   2) Строки без представительства НЕ выбрасываются, а попадают в «(без представительства)» —
//      иначе итог по представительствам не сходится с итогом компании.
//   3) Если в файле нет столбца представительства (например, «Товар в пути»),
//      вся сумма идёт в «(без представительства)», итог компании сохраняется.
//
// Параметр: слово, которое должно быть в имени файла: "остатки" или "товар в пути".

(СловоВИмениФайла as text) as table =>
let
    ПапкаИсточник = "C:\Users\v.lis\Desktop\ФинЦикл\Расчёт и анализ ФЦ\Запасы",
    Источник = Folder.Files(ПапкаИсточник),
    БезВременных = Table.SelectRows(Источник, each not Text.StartsWith([Name], "~$")),
    Файлы = Table.SelectRows(БезВременных, each Text.Contains(Text.Lower([Name]), СловоВИмениФайла)
        and (Text.EndsWith(Text.Lower([Name]), ".xlsx") or Text.EndsWith(Text.Lower([Name]), ".xls"))),

    Норм = (t as any) as text => Text.Replace(Text.Lower(if t = null then "" else Text.From(t)), "ё", "е"),

    ВыбратьЛист = (bookContent as binary, словоШапки as text) as table =>
        let
            Книга = Excel.Workbook(bookContent, true),
            Проверить = (data as table) as nullable table =>
                let
                    ДостаточноСтолбцов = Table.ColumnCount(data) >= 5,
                    ЕстьСразу = ДостаточноСтолбцов and List.AnyTrue(List.Transform(Table.ColumnNames(data), each Text.Contains(Норм(_), словоШапки))),
                    ПослеПромо = if ЕстьСразу or not ДостаточноСтолбцов then null else (try Table.PromoteHeaders(data, [PromoteAllScalars = true]) otherwise null),
                    ЕстьПослеПромо = if ПослеПромо = null then false
                        else List.AnyTrue(List.Transform(Table.ColumnNames(ПослеПромо), each Text.Contains(Норм(_), словоШапки)))
                in
                    if ЕстьСразу then data
                    else if ЕстьПослеПромо then ПослеПромо
                    else null,
            СПроверкой = Table.AddColumn(Книга, "Checked", each try Проверить([Data]) otherwise null),
            Подходящие = Table.SelectRows(СПроверкой, each [Checked] <> null),
            Результат = if Table.RowCount(Подходящие) = 0
                then error ("Не найден лист со столбцом, содержащим: " & словоШапки)
                else Подходящие{0}[Checked]
        in
            Результат,

    НайтиСтолбец = (t as table, содержит as list) as text =>
        let
            Совпадения = List.Select(Table.ColumnNames(t), each
                let низ = Норм(_) in List.AllTrue(List.Transform(содержит, each Text.Contains(низ, _)))),
            Результат = if List.Count(Совпадения) > 0 then Совпадения{0}
                else error ("Столбец не найден: " & Text.Combine(содержит, "+") & ". Реальные столбцы: " & Text.Combine(Table.ColumnNames(t), " | "))
        in
            Результат,

    ЧитатьКнигу = (Содержимое as binary) as table =>
        let
            Лист = ВыбратьЛист(Содержимое, "дата"),
            СтДата = НайтиСтолбец(Лист, {"дата"}),
            СтСумма = НайтиСтолбец(Лист, {"сумма", "уц"}),
            СтПредст = try НайтиСтолбец(Лист, {"представительств"}) otherwise null,
            Выбрано = if СтПредст = null
                then Table.AddColumn(Table.SelectColumns(Лист, {СтДата, СтСумма}), "Представительство", each null)
                else Table.SelectColumns(Лист, {СтДата, СтСумма, СтПредст}),
            Переименовано = Table.RenameColumns(Выбрано, List.RemoveNulls({
                {СтДата, "Дата"}, {СтСумма, "СуммаТг"},
                if СтПредст = null then null else {СтПредст, "Представительство"}}))
        in
            Переименовано,

    СДанными = Table.AddColumn(Файлы, "Данные", each try ЧитатьКнигу([Content]) otherwise null),
    Успешные = Table.SelectRows(СДанными, each [Данные] <> null),
    Нужное = Table.SelectColumns(Успешные, {"Name", "Date modified", "Данные"}),
    Раскрыто = Table.ExpandTableColumn(Нужное, "Данные", {"Дата", "СуммаТг", "Представительство"}),
    Типы = Table.TransformColumnTypes(Раскрыто, {{"Дата", type date}, {"СуммаТг", type number}}),
    БезОшибок = Table.RemoveRowsWithErrors(Типы, {"Дата", "СуммаТг"}),
    КонецМесяца = Table.SelectRows(БезОшибок, each [Дата] <> null and [СуммаТг] <> null
        and Date.EndOfMonth([Дата]) = [Дата]),

    // один месяц = один файл: берём файл с самой поздней датой изменения
    ЛучшийФайл = Table.Group(КонецМесяца, {"Дата"}, {{"ЛучшийФайл", each
        Table.First(Table.Sort(Table.Distinct(Table.SelectColumns(_, {"Name", "Date modified"})),
            {{"Date modified", Order.Descending}}))[Name], type text}}),
    Соединено = Table.NestedJoin(КонецМесяца, {"Дата"}, ЛучшийФайл, {"Дата"}, "Л", JoinKind.Inner),
    СФайлом = Table.ExpandTableColumn(Соединено, "Л", {"ЛучшийФайл"}),
    ОдинФайл = Table.SelectRows(СФайлом, each [Name] = [ЛучшийФайл]),

    ИменаЧистые = Table.TransformColumns(ОдинФайл, {{"Представительство", each
        if _ = null or Text.Trim(Text.From(_)) = "" then "(без представительства)" else Text.Trim(Text.From(_)), type text}}),
    Итог = Table.Group(ИменаЧистые, {"Дата", "Представительство"}, {{"СуммаТг", each List.Sum([СуммаТг]), type number}}),
    Результат = Table.RenameColumns(Итог, {{"Дата", "Период"}})
in
    Результат
