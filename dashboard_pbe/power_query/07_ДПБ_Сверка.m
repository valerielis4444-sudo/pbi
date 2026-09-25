// ДПБ_Сверка — проверка после каждого обновления.
//  Новые месяцы: сумма ДЗ по ПБЕ = ДЗ покупателей в общей модели (запрос «ДЗ и КЗ» -> лист «Данные»).
//  Исторические месяцы: проверка чтения файлов — оборот по дебету из файлов против истории
//  (подтверждает, что новые месяцы считаются так же, как история).
let
    Посл = List.Max(ДПБ_История[Период]),
    Факт = Table.Buffer(ДПБ_Факт),
    ФактМес = Table.Group(Факт, {"Период"}, {{"ДЗ_дашборд", each List.Sum([ДЗТг]), type number}}),
    Модель = Table.SelectRows(#"ДЗ и КЗ", each [Период] = Date.EndOfMonth([Период]) and [Период] > Посл),
    С1 = Table.ExpandTableColumn(Table.NestedJoin(ФактМес, {"Период"}, Модель, {"Период"}, "М", JoinKind.LeftOuter), "М", {"ДЗ_покупатели"}, {"ДЗ_модель"}),
    Файлы = Table.Group(ДПБ_Файлы, {"Период"}, {{"Дт_файлы", each List.Sum([ОборотДтТг]), type number}}),
    Ист = Table.Group(ДПБ_История, {"Период"}, {{"Дт_история", each List.Sum([ОборотДтТг]), type number}}),
    С2 = Table.ExpandTableColumn(Table.NestedJoin(С1, {"Период"}, Файлы, {"Период"}, "Ф", JoinKind.LeftOuter), "Ф", {"Дт_файлы"}),
    С3 = Table.ExpandTableColumn(Table.NestedJoin(С2, {"Период"}, Ист, {"Период"}, "И", JoinKind.LeftOuter), "И", {"Дт_история"}),
    Н = (x) => if x = null then 0 else x,
    Д1 = Table.AddColumn(С3, "Расхождение_ДЗ", each if [Период] > Посл then Н([ДЗ_дашборд]) - Н([ДЗ_модель]) else null, type nullable number),
    Д2 = Table.AddColumn(Д1, "Файлы_к_истории_Дт", each
        if [Период] <= Посл and [Дт_файлы] <> null and [Дт_история] <> null and [Дт_история] <> 0
        then [Дт_файлы] / [Дт_история] - 1 else null, type nullable number),
    Статус = Table.AddColumn(Д2, "Статус", each
        if [Период] > Посл then (if Number.Abs([Расхождение_ДЗ]) <= 1 then "OK" else "РАСХОЖДЕНИЕ")
        else if [Файлы_к_истории_Дт] = null then "история"
        else if Number.Abs([Файлы_к_истории_Дт]) <= 0.01 then "история: файлы сходятся"
        else "история: файлы НЕ сходятся", type text),
    Сорт = Table.Sort(Статус, {{"Период", Order.Descending}})
in
    Сорт
