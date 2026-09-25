// ДПБ_Календарь — месяцы от первого до последнего в ДПБ_Факт
let
    Даты = List.Buffer(List.Distinct(ДПБ_Факт[Период])),
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
    Доб4 = Table.AddColumn(Доб3, "МесяцИндекс", each [Год] * 12 + [НомерМесяца], Int64.Type),
    Доб5 = Table.AddColumn(Доб4, "Дней", each Date.Day([Дата]), Int64.Type)
in
    Доб5
