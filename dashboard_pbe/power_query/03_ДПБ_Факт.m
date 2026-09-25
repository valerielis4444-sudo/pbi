// ДПБ_Факт — история + новые месяцы из папки (всё, что позже последней даты истории)
let
    История = Table.Buffer(ДПБ_История),
    Посл = List.Max(История[Период]),
    Новые = Table.SelectRows(ДПБ_Файлы, each [Период] > Посл),
    Все = Table.Combine({История, Новые}),
    БезПустых = Table.SelectRows(Все, each [ДЗТг] <> 0 or [АвансыТг] <> 0 or [ОборотДтТг] <> 0 or [ОборотКтТг] <> 0),
    Типы = Table.TransformColumnTypes(БезПустых, {{"Период", type date}, {"ПБЕ", type text},
        {"ДЗТг", type number}, {"АвансыТг", type number}, {"ОборотДтТг", type number}, {"ОборотКтТг", type number}})
in
    Типы
