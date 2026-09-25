// Имя запроса: ДП_Продажи
// Загрузка: «Только создать подключение»
//
// Выручка и себестоимость по представительствам из куба GEFEST/Sales_V1.
// Отличие от старого «Продажи_Представительства»: продажи без представительства
// НЕ выбрасываются, а идут в «(без представительства)» — иначе итог не сходится с кубом
// (в августе 2026 старый запрос терял 6,7 млн ₸ выручки).
// Берутся только закрытые месяцы (текущий месяц отсекается).

let
    Источник = AnalysisServices.Database("GEFEST", "Sales_V1", [Culture = "en-US"]),
    Sales = Источник{[Id = "Sales"]}[Data],
    Sales1 = Sales{[Id = "Sales"]}[Data],
    Куб = Cube.Transform(Sales1,
        {
            {Cube.AddAndExpandDimensionColumn, "[Период]", {"[Период].[Дата].[Год]", "[Период].[Дата].[Месяц]"}, {"Период.Год", "Период.Месяц"}},
            {Cube.AddAndExpandDimensionColumn, "[Номенклатор]", {"[Номенклатор].[Представительство].[Представительство]"}, {"Представительство"}},
            {Cube.AddMeasureColumn, "Сумма без НДС", "[Measures].[Сумма без НДС]"},
            {Cube.AddMeasureColumn, "Сумма в уч. ценах", "[Measures].[Сумма в уч. ценах]"}
        }),
    Месяцы = {"Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"},
    ТолькоМесяцы = Table.SelectRows(Куб, each List.Contains(Месяцы, [Период.Месяц])),
    ГодВЧисло = (y as any) as number => try Number.From(y) otherwise Number.FromText(Text.From(y)),
    ДобПериод = Table.AddColumn(ТолькоМесяцы, "Период", each
        Date.EndOfMonth(#date(ГодВЧисло([Период.Год]), List.PositionOf(Месяцы, [Период.Месяц]) + 1, 1)), type date),
    Закрытые = Table.SelectRows(ДобПериод, each [Период] < Date.StartOfMonth(DateTime.Date(DateTime.LocalNow()))),
    ИменаЧистые = Table.TransformColumns(Закрытые, {{"Представительство", each
        if _ = null or Text.Trim(Text.From(_)) = "" then "(без представительства)" else Text.Trim(Text.From(_)), type text}}),
    Итог = Table.Group(ИменаЧистые, {"Период", "Представительство"}, {
        {"ВыручкаТг", each List.Sum([Сумма без НДС]), type number},
        {"СебестоимостьТг", each List.Sum([Сумма в уч. ценах]), type number}})
in
    Итог
