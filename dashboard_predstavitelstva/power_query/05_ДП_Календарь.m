// Имя запроса: ДП_Календарь
// Загрузка: «Только создать подключение» + «Добавить эти данные в модель данных»
//
// Одна строка = один месяц (дата конца месяца), без пропусков, от первого до последнего месяца в ДП_Факт.
// «Месяц» записан как «2026-08 авг» — так он сортируется по времени, а не по алфавиту.
// «МесяцИндекс» нужен мерам «предыдущий период» и «год назад».

let
    Даты = List.Distinct(ДП_Факт[Период]),
    Мин = List.Min(Даты),
    Макс = List.Max(Даты),
    Кол = (Date.Year(Макс) - Date.Year(Мин)) * 12 + Date.Month(Макс) - Date.Month(Мин) + 1,
    Месяцы = List.Transform({0 .. Кол - 1}, each Date.EndOfMonth(Date.AddMonths(Date.StartOfMonth(Мин), _))),
    Таблица = Table.FromList(Месяцы, Splitter.SplitByNothing(), {"Дата"}),
    Типы = Table.TransformColumnTypes(Таблица, {{"Дата", type date}}),
    НазванияМес = {"янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"},
    Доб1 = Table.AddColumn(Типы, "Год", each Date.Year([Дата]), Int64.Type),
    Доб2 = Table.AddColumn(Доб1, "НомерМесяца", each Date.Month([Дата]), Int64.Type),
    Доб3 = Table.AddColumn(Доб2, "Месяц", each Text.From([Год]) & "-" & Text.PadStart(Text.From([НомерМесяца]), 2, "0")
        & " " & НазванияМес{[НомерМесяца] - 1}, type text),
    Доб4 = Table.AddColumn(Доб3, "Квартал", each Text.From([Год]) & " Q" & Text.From(Date.QuarterOfYear([Дата])), type text),
    Доб5 = Table.AddColumn(Доб4, "МесяцИндекс", each [Год] * 12 + [НомерМесяца], Int64.Type),
    Доб6 = Table.AddColumn(Доб5, "Дней", each Date.Day([Дата]), Int64.Type)
in
    Доб6
