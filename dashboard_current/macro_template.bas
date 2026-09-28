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

%%QFUNCS%%

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
    For Each qn In Array("Тек_Итог", "Тек_Продажи", "fnТек_Остатки", "Тек_Параметры")
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
    ws.Columns("D:M").ColumnWidth = 14
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
    wb.Queries.Add "fnТек_Остатки", QM1()
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
    Chk "загрузка таблицы (если ошибка про остатки — нет файла остатков на отчётную дату)"
    Set lo = ws.ListObjects("Тек_Итог")
    lo.TableStyle = "TableStyleMedium2"
    fmts = Array("", "", "#,##0.0", "#,##0.0", "#,##0.0", "0", "0.0", "#,##0.0", "#,##0.0", "0", "0.0", "+0.0;-0.0;0.0")
    For i = 1 To lo.ListColumns.Count
        If i <= 12 Then If fmts(i - 1) <> "" Then lo.ListColumns(i).Range.NumberFormat = fmts(i - 1)
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
    ws.Range("B9").Value = "Остаток = склад + товар в пути, в учётных ценах. Итоги считаются по видимым строкам (учитывают фильтр таблицы)."
    ws.Range("B8:B9").WrapText = True
    ws.Range("B8:B9").Font.Size = 9
    ws.Rows("8:9").RowHeight = 30
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
