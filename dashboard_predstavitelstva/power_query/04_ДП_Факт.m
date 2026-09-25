// Имя запроса: ДП_Факт
// Загрузка: «Только создать подключение» + галочка «Добавить эти данные в модель данных»
//
// Единая таблица фактов: строка = представительство × месяц.
//   История (янв.2024 — последний месяц истории) — с листа « Представительства (историч.)»
//   Новые месяцы — из куба и папки «Запасы».
// Граница определяется автоматически: всё, что позже последней даты истории, берётся из источников.
// Новый месяц попадает в дашборд, только когда есть И продажи, И остатки на складах —
// чтобы не было месяца с нулевыми запасами и DIO = 0.

let
    // Table.Buffer — каждый источник (лист, куб, папка) читается ОДИН раз
    История = Table.Buffer(ДП_История),
    ПоследняяДатаИстории = List.Max(История[Период]),

    Продажи = Table.Buffer(Table.SelectRows(ДП_Продажи, each [Период] > ПоследняяДатаИстории)),
    Склад = Table.Buffer(Table.SelectRows(fnДП_Остатки("остатки"), each [Период] > ПоследняяДатаИстории)),
    Путь = Table.Buffer(Table.SelectRows(fnДП_Остатки("товар в пути"), each [Период] > ПоследняяДатаИстории)),

    ПолныеМесяцы = List.Buffer(List.Intersect({List.Distinct(Продажи[Период]), List.Distinct(Склад[Период])})),

    П = Table.AddColumn(Table.AddColumn(Продажи, "ЗапасыСкладТг", each 0, type number), "ТоварВПутиТг", each 0, type number),
    С0 = Table.RenameColumns(Склад, {{"СуммаТг", "ЗапасыСкладТг"}}),
    С = Table.AddColumn(Table.AddColumn(Table.AddColumn(С0, "ВыручкаТг", each 0, type number), "СебестоимостьТг", each 0, type number), "ТоварВПутиТг", each 0, type number),
    В0 = Table.RenameColumns(Путь, {{"СуммаТг", "ТоварВПутиТг"}}),
    В = Table.AddColumn(Table.AddColumn(Table.AddColumn(В0, "ВыручкаТг", each 0, type number), "СебестоимостьТг", each 0, type number), "ЗапасыСкладТг", each 0, type number),

    НовыеВсе = Table.Group(Table.Combine({П, С, В}), {"Период", "Представительство"}, {
        {"ВыручкаТг", each List.Sum([ВыручкаТг]), type number},
        {"СебестоимостьТг", each List.Sum([СебестоимостьТг]), type number},
        {"ЗапасыСкладТг", each List.Sum([ЗапасыСкладТг]), type number},
        {"ТоварВПутиТг", each List.Sum([ТоварВПутиТг]), type number}}),
    Новые = Table.SelectRows(НовыеВсе, each List.Contains(ПолныеМесяцы, [Период])),

    Все = Table.Combine({История, Новые}),
    СЗапасами = Table.AddColumn(Все, "ЗапасыТг", each [ЗапасыСкладТг] + [ТоварВПутиТг], type number),
    БезПустыхСтрок = Table.SelectRows(СЗапасами, each [ВыручкаТг] <> 0 or [СебестоимостьТг] <> 0 or [ЗапасыТг] <> 0),
    Типы = Table.TransformColumnTypes(БезПустыхСтрок, {
        {"Период", type date}, {"Представительство", type text},
        {"ВыручкаТг", type number}, {"СебестоимостьТг", type number},
        {"ЗапасыСкладТг", type number}, {"ТоварВПутиТг", type number}, {"ЗапасыТг", type number}})
in
    Типы
