' =====================================================================
'  Лист «Текущая_оборачиваемость» — DIO по представительствам на дату внутри месяца
'  (остаток на отчётную дату, продажи/себестоимость с 1-го числа, дней = число даты).
'  Запуск: Alt+F11 -> Insert -> Module -> вставить -> курсор в BuildCurrent -> F5.
'  Потом отчётную дату меняете в ячейке B5 и жмёте «Данные -> Обновить всё».
' =====================================================================
Option Explicit
Private wb As Workbook
Private LogText As String

Private Sub Chk(ByVal s As String)
    If Err.Number <> 0 Then LogText = LogText & "- " & s & ": " & Err.Description & vbLf: Err.Clear
End Sub

Private Function QM0() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Источник = Excel.CurrentWorkbook(){[Name = ""tblТек_Параметры""]}[Content]," & vbLf
    s = s & "    Дата = Date.From(Источник{0}[Дата_отчета])" & vbLf
    s = s & "in" & vbLf
    s = s & "    Дата" & vbLf
    QM0 = s
End Function

Private Function QM1() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    ДатаОтчета = Тек_Параметры," & vbLf
    s = s & "    Источник = AnalysisServices.Database(""GEFEST"", ""Remainder_V1"", [Culture = ""en-US""])," & vbLf
    s = s & "    Модель = Источник{[Id = ""Модель""]}[Data]," & vbLf
    s = s & "    Модель1 = Модель{[Id = ""Модель""]}[Data]," & vbLf
    s = s & "    Куб = Cube.Transform(Модель1," & vbLf
    s = s & "        {" & vbLf
    s = s & "            {Cube.AddAndExpandDimensionColumn, ""[Номенклатор]"", {""[Номенклатор].[Представительство].[Представительство]""}, {""Представительство""}}," & vbLf
    s = s & "            {Cube.AddAndExpandDimensionColumn, ""[Период]"", {""[Период].[Дата].[Дата]""}, {""ДатаТекст""}}," & vbLf
    s = s & "            {Cube.AddMeasureColumn, ""Количество"", ""[Measures].[Количество]""}," & vbLf
    s = s & "            {Cube.AddMeasureColumn, ""Сумма в уч. ценах"", ""[Measures].[Сумма в уч. ценах]""}" & vbLf
    s = s & "        })," & vbLf
    s = s & "    СДатой = Table.AddColumn(Куб, ""Дата"", each" & vbLf
    s = s & "        try Date.From([ДатаТекст]) otherwise (try Date.FromText(Text.From([ДатаТекст]), ""en-US"") otherwise null), type nullable date)," & vbLf
    s = s & "    Буфер = Table.Buffer(СДатой)," & vbLf
    s = s & "    ДатыВКубе = List.Sort(List.Distinct(List.RemoveNulls(Буфер[Дата])))," & vbLf
    s = s & "    Отбор = Table.SelectRows(Буфер, each [Дата] = ДатаОтчета)," & vbLf
    s = s & "    Проверено = if Table.RowCount(Отбор) = 0" & vbLf
    s = s & "        then error (""В кубе остатков нет даты "" & Date.ToText(ДатаОтчета) & "". Есть даты: """ & vbLf
    s = s & "            & Text.Combine(List.Transform(ДатыВКубе, Date.ToText), "", "") & "". Поставьте одну из них в ячейку B5."")" & vbLf
    s = s & "        else Отбор," & vbLf
    s = s & "    ИменаЧистые = Table.TransformColumns(Проверено, {{""Представительство"", each" & vbLf
    s = s & "        if _ = null or Text.Trim(Text.From(_)) = """" then ""(без представительства)"" else Text.Trim(Text.From(_)), type text}})," & vbLf
    s = s & "    Итог = Table.Group(ИменаЧистые, {""Представительство""}, {" & vbLf
    s = s & "        {""СуммаТг"", each List.Sum([Сумма в уч. ценах]), type number}," & vbLf
    s = s & "        {""КоличествоШт"", each List.Sum([Количество]), type number}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Итог" & vbLf
    QM1 = s
End Function

Private Function QM2() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Дата = Тек_Параметры," & vbLf
    s = s & "    Период = Date.EndOfMonth(Дата)," & vbLf
    s = s & "    Источник = AnalysisServices.Database(""GEFEST"", ""Sales_V1"", [Culture = ""en-US""])," & vbLf
    s = s & "    Sales = Источник{[Id = ""Sales""]}[Data]," & vbLf
    s = s & "    Sales1 = Sales{[Id = ""Sales""]}[Data]," & vbLf
    s = s & "    Куб = Cube.Transform(Sales1," & vbLf
    s = s & "        {" & vbLf
    s = s & "            {Cube.AddAndExpandDimensionColumn, ""[Период]"", {""[Период].[Дата].[Год]"", ""[Период].[Дата].[Месяц]""}, {""Период.Год"", ""Период.Месяц""}}," & vbLf
    s = s & "            {Cube.AddAndExpandDimensionColumn, ""[Номенклатор]"", {""[Номенклатор].[Представительство].[Представительство]""}, {""Представительство""}}," & vbLf
    s = s & "            {Cube.AddMeasureColumn, ""Сумма без НДС"", ""[Measures].[Сумма без НДС]""}," & vbLf
    s = s & "            {Cube.AddMeasureColumn, ""Сумма в уч. ценах"", ""[Measures].[Сумма в уч. ценах]""}" & vbLf
    s = s & "        })," & vbLf
    s = s & "    Месяцы = {""Январь"", ""Февраль"", ""Март"", ""Апрель"", ""Май"", ""Июнь"", ""Июль"", ""Август"", ""Сентябрь"", ""Октябрь"", ""Ноябрь"", ""Декабрь""}," & vbLf
    s = s & "    ТолькоМесяцы = Table.SelectRows(Куб, each List.Contains(Месяцы, [Период.Месяц]))," & vbLf
    s = s & "    ГодВЧисло = (y as any) as number => try Number.From(y) otherwise Number.FromText(Text.From(y))," & vbLf
    s = s & "    ДобПериод = Table.AddColumn(ТолькоМесяцы, ""Период"", each" & vbLf
    s = s & "        Date.EndOfMonth(#date(ГодВЧисло([Период.Год]), List.PositionOf(Месяцы, [Период.Месяц]) + 1, 1)), type date)," & vbLf
    s = s & "    ЭтотМесяц = Table.SelectRows(ДобПериод, each [Период] = Период)," & vbLf
    s = s & "    ИменаЧистые = Table.TransformColumns(ЭтотМесяц, {{""Представительство"", each" & vbLf
    s = s & "        if _ = null or Text.Trim(Text.From(_)) = """" then ""(без представительства)"" else Text.Trim(Text.From(_)), type text}})," & vbLf
    s = s & "    Итог = Table.Group(ИменаЧистые, {""Представительство""}, {" & vbLf
    s = s & "        {""ВыручкаТг"", each List.Sum([Сумма без НДС]), type number}," & vbLf
    s = s & "        {""СебестоимостьТг"", each List.Sum([Сумма в уч. ценах]), type number}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Итог" & vbLf
    QM2 = s
End Function

Private Function QM3() As String
    Dim s As String
    s = s & "let" & vbLf
    s = s & "    Дата = Тек_Параметры," & vbLf
    s = s & "    ДнейМес = Date.Day(Дата)," & vbLf
    s = s & "    ПредМес = Date.EndOfMonth(Date.AddMonths(Date.StartOfMonth(Дата), -1))," & vbLf
    s = s & "    ДнейПМ = Date.Day(ПредМес)," & vbLf
    s = s & "" & vbLf
    s = s & "    Пр = Тек_Продажи," & vbLf
    s = s & "    Скл = Тек_Остатки," & vbLf
    s = s & "    Ф = Table.SelectRows(ДП_Факт, each [Период] = ПредМес)," & vbLf
    s = s & "" & vbLf
    s = s & "    Доб = (t as table, имена as list) as table =>" & vbLf
    s = s & "        List.Accumulate(имена, t, (s, n) => if Table.HasColumns(s, n) then s else Table.AddColumn(s, n, each 0, type number))," & vbLf
    s = s & "    Все = {""ВыручкаТг"", ""СебестоимостьТг"", ""ОстатокТг"", ""ОстатокШт"", ""СебПредТг"", ""ЗапПредТг""}," & vbLf
    s = s & "    A = Доб(Table.SelectColumns(Пр, {""Представительство"", ""ВыручкаТг"", ""СебестоимостьТг""}), Все)," & vbLf
    s = s & "    B = Доб(Table.RenameColumns(Table.SelectColumns(Скл, {""Представительство"", ""СуммаТг"", ""КоличествоШт""})," & vbLf
    s = s & "            {{""СуммаТг"", ""ОстатокТг""}, {""КоличествоШт"", ""ОстатокШт""}}), Все)," & vbLf
    s = s & "    D = Доб(Table.RenameColumns(Table.SelectColumns(Ф, {""Представительство"", ""СебестоимостьТг"", ""ЗапасыСкладТг""})," & vbLf
    s = s & "            {{""СебестоимостьТг"", ""СебПредТг""}, {""ЗапасыСкладТг"", ""ЗапПредТг""}}), Все)," & vbLf
    s = s & "    Сумма = Table.Group(Table.Combine({A, B, D}), {""Представительство""}, {" & vbLf
    s = s & "        {""ВыручкаТг"", each List.Sum([ВыручкаТг]), type number}," & vbLf
    s = s & "        {""СебестоимостьТг"", each List.Sum([СебестоимостьТг]), type number}," & vbLf
    s = s & "        {""ОстатокТг"", each List.Sum([ОстатокТг]), type number}," & vbLf
    s = s & "        {""ОстатокШт"", each List.Sum([ОстатокШт]), type number}," & vbLf
    s = s & "        {""СебПредТг"", each List.Sum([СебПредТг]), type number}," & vbLf
    s = s & "        {""ЗапПредТг"", each List.Sum([ЗапПредТг]), type number}})," & vbLf
    s = s & "    Непустые = Table.SelectRows(Сумма, each [ВыручкаТг] <> 0 or [СебестоимостьТг] <> 0 or [ОстатокТг] <> 0)," & vbLf
    s = s & "" & vbLf
    s = s & "    Классы = Table.SelectColumns(ДП_Представительства, {""Представительство"", ""Класс ABC""})," & vbLf
    s = s & "    СКлассом = Table.ExpandTableColumn(Table.NestedJoin(Непустые, {""Представительство""}, Классы, {""Представительство""}, ""К"", JoinKind.LeftOuter), ""К"", {""Класс ABC""})," & vbLf
    s = s & "" & vbLf
    s = s & "    Р = Table.AddColumn(СКлассом, ""Итог"", each [" & vbLf
    s = s & "        #""Выручка с 1-го числа, млн тг"" = [ВыручкаТг] / 1000000," & vbLf
    s = s & "        #""Себестоимость с 1-го числа, млн тг"" = [СебестоимостьТг] / 1000000," & vbLf
    s = s & "        #""Остаток на дату, млн тг"" = [ОстатокТг] / 1000000," & vbLf
    s = s & "        #""Остаток на дату, шт."" = [ОстатокШт]," & vbLf
    s = s & "        #""Дней"" = ДнейМес," & vbLf
    s = s & "        #""DIO на дату"" = if [СебестоимостьТг] > 0 then [ОстатокТг] * ДнейМес / [СебестоимостьТг] else null," & vbLf
    s = s & "        #""Себестоимость пред. месяца, млн тг"" = [СебПредТг] / 1000000," & vbLf
    s = s & "        #""Остаток на конец пред. месяца, млн тг"" = [ЗапПредТг] / 1000000," & vbLf
    s = s & "        #""Дней пред. месяца"" = ДнейПМ," & vbLf
    s = s & "        #""DIO пред. месяца"" = if [СебПредТг] > 0 then [ЗапПредТг] * ДнейПМ / [СебПредТг] else null" & vbLf
    s = s & "    ])," & vbLf
    s = s & "    Раскрыто = Table.ExpandRecordColumn(Р, ""Итог"", {""Выручка с 1-го числа, млн тг"", ""Себестоимость с 1-го числа, млн тг""," & vbLf
    s = s & "        ""Остаток на дату, млн тг"", ""Остаток на дату, шт."", ""Дней"", ""DIO на дату"", ""Себестоимость пред. месяца, млн тг""," & vbLf
    s = s & "        ""Остаток на конец пред. месяца, млн тг"", ""Дней пред. месяца"", ""DIO пред. месяца""})," & vbLf
    s = s & "    Изм = Table.AddColumn(Раскрыто, ""Изменение DIO"", each" & vbLf
    s = s & "        if [DIO на дату] <> null and [DIO пред. месяца] <> null then [DIO на дату] - [DIO пред. месяца] else null)," & vbLf
    s = s & "    Выбор = Table.SelectColumns(Изм, {""Представительство"", ""Класс ABC"", ""Выручка с 1-го числа, млн тг"", ""Себестоимость с 1-го числа, млн тг""," & vbLf
    s = s & "        ""Остаток на дату, млн тг"", ""Остаток на дату, шт."", ""Дней"", ""DIO на дату"", ""Себестоимость пред. месяца, млн тг""," & vbLf
    s = s & "        ""Остаток на конец пред. месяца, млн тг"", ""Дней пред. месяца"", ""DIO пред. месяца"", ""Изменение DIO""})," & vbLf
    s = s & "    Типы = Table.TransformColumnTypes(Выбор, {{""Представительство"", type text}, {""Класс ABC"", type text}," & vbLf
    s = s & "        {""Выручка с 1-го числа, млн тг"", type number}, {""Себестоимость с 1-го числа, млн тг"", type number}," & vbLf
    s = s & "        {""Остаток на дату, млн тг"", type number}, {""Остаток на дату, шт."", type number}, {""Дней"", Int64.Type}, {""DIO на дату"", type number}," & vbLf
    s = s & "        {""Себестоимость пред. месяца, млн тг"", type number}, {""Остаток на конец пред. месяца, млн тг"", type number}," & vbLf
    s = s & "        {""Дней пред. месяца"", Int64.Type}, {""DIO пред. месяца"", type number}, {""Изменение DIO"", type number}})," & vbLf
    s = s & "    Сорт = Table.Sort(Типы, {{""Выручка с 1-го числа, млн тг"", Order.Descending}})" & vbLf
    s = s & "in" & vbLf
    s = s & "    Сорт" & vbLf
    QM3 = s
End Function


Public Sub BuildCurrent()
    Dim ws As Worksheet, lo As ListObject, i As Long, c, labels, frm, fmts, r As Range, qn, TG As String
    Set wb = ThisWorkbook
    LogText = ""
    TG = ChrW(8376)
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Текущая_оборачиваемость").Delete
    For i = wb.Connections.Count To 1 Step -1
        If InStr(wb.Connections(i).Name, "Тек_") > 0 Then wb.Connections(i).Delete
    Next i
    For Each qn In Array("Тек_Итог", "Тек_Продажи", "Тек_Остатки", "fnТек_Остатки", "Тек_Параметры")
        wb.Queries(qn).Delete
    Next qn
    Err.Clear
    Application.DisplayAlerts = True

    ' ---------- лист и параметр ----------
    If Not wb.Worksheets("Карта_DIO") Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Карта_DIO"))
    End If
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    Err.Clear
    ws.Name = "Текущая_оборачиваемость"
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D:N").ColumnWidth = 14
    ws.Range("B2").Value = "ТЕКУЩАЯ ОБОРАЧИВАЕМОСТЬ ЗАПАСОВ ПО ПРЕДСТАВИТЕЛЬСТВАМ — МЕСЯЦ НА ДАТУ"
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(31, 56, 100)
    ws.Range("B4").Value = "Дата_отчета"
    ws.Range("B5").Value = DateSerial(2026, 9, 27)
    ws.Range("B5").NumberFormat = "dd.mm.yyyy"
    ws.ListObjects.Add(xlSrcRange, ws.Range("B4:B5"), , xlYes).Name = "tblТек_Параметры"
    ws.Range("B5").Font.Size = 14
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").Interior.Color = RGB(255, 242, 204)
    ws.Range("C5").Value = "<- отчётная дата: остатки на эту дату, продажи с 1-го числа месяца. Поменяйте дату и нажмите «Данные -> Обновить всё»."
    ws.Range("C5").Font.Italic = True
    ws.Range("C5").Font.Color = RGB(120, 120, 120)
    Chk "лист и параметр"

    ' ---------- запросы ----------
    wb.Queries.Add "Тек_Параметры", QM0()
    wb.Queries.Add "Тек_Остатки", QM1()
    wb.Queries.Add "Тек_Продажи", QM2()
    wb.Queries.Add "Тек_Итог", QM3()
    Chk "создание запросов"

    ' ---------- таблица ----------
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Тек_Итог;Extended Properties=""""", _
        Destination:=ws.Range("$B$12"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [Тек_Итог]")
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = False
        .ListObject.DisplayName = "Тек_Итог"
        .Refresh BackgroundQuery:=False
    End With
    Chk "загрузка таблицы (если в тексте ошибки «нет даты» — поставьте в B5 дату, которая есть в кубе остатков)"
    Set lo = ws.ListObjects("Тек_Итог")
    lo.TableStyle = "TableStyleMedium2"
    fmts = Array("", "", "#,##0.0", "#,##0.0", "#,##0.0", "#,##0", "0", "0.0", "#,##0.0", "#,##0.0", "0", "0.0", "+0.0;-0.0;0.0")
    For i = 1 To lo.ListColumns.Count
        If i <= 13 Then If fmts(i - 1) <> "" Then lo.ListColumns(i).Range.NumberFormat = fmts(i - 1)
    Next i
    lo.HeaderRowRange.WrapText = True
    lo.HeaderRowRange.RowHeight = 45
    lo.HeaderRowRange.VerticalAlignment = xlCenter
    Set r = lo.ListColumns("Выручка с 1-го числа, млн тг").DataBodyRange
    With r.FormatConditions.AddDatabar
        .BarColor.Color = RGB(46, 117, 182)
        .BarFillType = xlDataBarFillSolid
    End With
    Set r = lo.ListColumns("DIO на дату").DataBodyRange
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(99, 190, 123)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).Type = xlConditionValuePercentile
        .ColorScaleCriteria(3).Value = 95
        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
    End With
    Set r = lo.ListColumns("Изменение DIO").DataBodyRange
    r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=" & CStr(0.05), Formula2:="=100000").Font.Color = RGB(198, 40, 40)
    r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=-100000", Formula2:="=" & CStr(-0.05)).Font.Color = RGB(46, 125, 50)
    Chk "оформление таблицы"

    ' ---------- карточки итогов (по видимым строкам — учитывают фильтр таблицы) ----------
    c = Array("D", "F", "H", "J", "L")
    labels = Array("Выручка с 1-го числа, млн " & TG, "Себестоимость с 1-го числа, млн " & TG, "Остаток на дату, млн " & TG, _
                   "DIO на дату, дней", "DIO прошлого месяца, дней")
    frm = Array("=SUBTOTAL(109,Тек_Итог[Выручка с 1-го числа, млн тг])", _
                "=SUBTOTAL(109,Тек_Итог[Себестоимость с 1-го числа, млн тг])", _
                "=SUBTOTAL(109,Тек_Итог[Остаток на дату, млн тг])", _
                "=IFERROR(H8*DAY($B$5)/F8,""-"")", _
                "=IFERROR(SUBTOTAL(109,Тек_Итог[Остаток на конец пред. месяца, млн тг])*DAY(EOMONTH($B$5,-1))/SUBTOTAL(109,Тек_Итог[Себестоимость пред. месяца, млн тг]),""-"")")
    For i = 0 To 4
        With ws.Range(c(i) & "7:" & Chr(Asc(c(i)) + 1) & "9")
            .Interior.Color = RGB(242, 244, 248)
            .BorderAround xlContinuous, xlThin, , RGB(208, 215, 226)
        End With
        ws.Range(c(i) & "7").Value = labels(i)
        ws.Range(c(i) & "7").Font.Size = 9
        ws.Range(c(i) & "7").Font.Color = RGB(89, 89, 89)
        With ws.Range(c(i) & "8:" & Chr(Asc(c(i)) + 1) & "8")
            .Merge
            .HorizontalAlignment = xlCenter
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(31, 56, 100)
            .NumberFormat = IIf(i >= 3, "0.0", "#,##0.0")
        End With
        ws.Range(c(i) & "8").Formula = frm(i)
        ws.Range(c(i) & "9:" & Chr(Asc(c(i)) + 1) & "9").Merge
        ws.Range(c(i) & "9").HorizontalAlignment = xlCenter
        ws.Range(c(i) & "9").Font.Size = 9
    Next i
    ws.Range("J9").Formula = "=""дней в периоде: ""&DAY($B$5)"
    ws.Range("L9").Formula = "=IFERROR(J8-L8,"""")"
    ws.Range("L9").NumberFormat = "[Red]+0.0"" дн. к прошлому месяцу"";[Color10]-0.0"" дн. к прошлому месяцу"";0.0"
    ws.Range("B7").Value = "Как считается"
    ws.Range("B7").Font.Bold = True
    ws.Range("B8").Value = "DIO на дату = остаток на дату " & ChrW(215) & " число дней с 1-го числа " & ChrW(247) & " себестоимость с 1-го числа."
    ws.Range("B9").Value = "Остаток на дату — из OLAP (куб Remainder_V1), сумма в учётных ценах; штуки — для справки, в расчёт не входят. Прошлый месяц — остатки на складах на конец месяца. Итоги — по видимым строкам."
    ws.Range("B8:B9").WrapText = True
    ws.Range("B8:B9").Font.Size = 9
    ws.Rows("8:9").RowHeight = 45
    Chk "карточки"

    ws.Activate
    ws.Range("B13").Select
    ActiveWindow.FreezePanes = False
    ws.Range("A13").Select
    ActiveWindow.FreezePanes = True
    ws.Range("B5").Select
    On Error GoTo 0
    If LogText = "" Then
        MsgBox "Готово! Лист «Текущая_оборачиваемость» создан." & vbLf & "Дата меняется в ячейке B5, затем «Данные -> Обновить всё».", vbInformation
    Else
        MsgBox "Лист создан, но есть проблемы. Пришлите этот текст:" & vbLf & vbLf & LogText, vbExclamation
    End If
End Sub
