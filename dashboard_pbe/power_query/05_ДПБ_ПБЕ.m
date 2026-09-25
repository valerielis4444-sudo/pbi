// ДПБ_ПБЕ — справочник ПБЕ для срезов: тип и оборот за последние 12 месяцев
let
    Факт = Table.Buffer(ДПБ_Факт),
    Посл = List.Max(Факт[Период]),
    Гр12 = Date.EndOfMonth(Date.AddMonths(Посл, -12)),
    Имена = Table.Distinct(Table.SelectColumns(Факт, {"ПБЕ"})),
    Об12 = Table.Group(Table.SelectRows(Факт, each [Период] > Гр12), {"ПБЕ"}, {{"Оборот_12м", each List.Sum([ОборотДтТг]), type number}}),
    Соед = Table.ExpandTableColumn(Table.NestedJoin(Имена, {"ПБЕ"}, Об12, {"ПБЕ"}, "т", JoinKind.LeftOuter), "т", {"Оборот_12м"}),
    Нули = Table.ReplaceValue(Соед, null, 0, Replacer.ReplaceValue, {"Оборот_12м"}),
    Тип = Table.AddColumn(Нули, "Тип", each if [ПБЕ] = "(без метки ПБЕ)"
        then "Розница и маркетплейсы" else "Филиалы (юрлица)", type text),
    Сорт = Table.Sort(Тип, {{"Оборот_12м", Order.Descending}})
in
    Сорт
