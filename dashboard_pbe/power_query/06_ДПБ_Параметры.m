// ДПБ_Параметры — таблица tblДПБ_Параметры на листе ДПБ_Своды: Цель_DSO (= Параметры!C21) и Порог_доли
let
    Источник = Excel.CurrentWorkbook(){[Name = "tblДПБ_Параметры"]}[Content],
    Типы = Table.TransformColumnTypes(Источник, {{"Цель_DSO", type number}, {"Порог_доли", type number}})
in
    Типы
