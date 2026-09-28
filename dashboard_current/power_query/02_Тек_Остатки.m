// Тек_Остатки — остатки по представительствам из OLAP (сервер GEFEST, база Remainder_V1, куб «Модель»)
// на отчётную дату из ячейки B5. Мера «Сумма в уч. ценах» (в DIO), «Количество» — для справки.
// Дата в кубе приходит текстом в американском формате (9/27/2026) — разбираем с культурой en-US.
let
    ДатаОтчета = Тек_Параметры,
    Источник = AnalysisServices.Database("GEFEST", "Remainder_V1", [Culture = "en-US"]),
    Модель = Источник{[Id = "Модель"]}[Data],
    Модель1 = Модель{[Id = "Модель"]}[Data],
    Куб = Cube.Transform(Модель1,
        {
            {Cube.AddAndExpandDimensionColumn, "[Номенклатор]", {"[Номенклатор].[Представительство].[Представительство]"}, {"Представительство"}},
            {Cube.AddAndExpandDimensionColumn, "[Период]", {"[Период].[Дата].[Дата]"}, {"ДатаТекст"}},
            {Cube.AddMeasureColumn, "Количество", "[Measures].[Количество]"},
            {Cube.AddMeasureColumn, "Сумма в уч. ценах", "[Measures].[Сумма в уч. ценах]"}
        }),
    СДатой = Table.AddColumn(Куб, "Дата", each
        try Date.From([ДатаТекст]) otherwise (try Date.FromText(Text.From([ДатаТекст]), "en-US") otherwise null), type nullable date),
    Буфер = Table.Buffer(СДатой),
    ДатыВКубе = List.Sort(List.Distinct(List.RemoveNulls(Буфер[Дата]))),
    Отбор = Table.SelectRows(Буфер, each [Дата] = ДатаОтчета),
    Проверено = if Table.RowCount(Отбор) = 0
        then error ("В кубе остатков нет даты " & Date.ToText(ДатаОтчета) & ". Есть даты: "
            & Text.Combine(List.Transform(ДатыВКубе, Date.ToText), ", ") & ". Поставьте одну из них в ячейку B5.")
        else Отбор,
    ИменаЧистые = Table.TransformColumns(Проверено, {{"Представительство", each
        if _ = null or Text.Trim(Text.From(_)) = "" then "(без представительства)" else Text.Trim(Text.From(_)), type text}}),
    Итог = Table.Group(ИменаЧистые, {"Представительство"}, {
        {"СуммаТг", each List.Sum([Сумма в уч. ценах]), type number},
        {"КоличествоШт", each List.Sum([Количество]), type number}})
in
    Итог
