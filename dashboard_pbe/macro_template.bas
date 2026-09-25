' =====================================================================
'  Дашборд ПБЕ — полная автоматическая сборка
'
'  Делает всё сам: именованный диапазон, таблицу параметров, 7 запросов Power Query,
'  загрузку в модель данных, связи, 31 меру DAX, служебные сводные, лист «Дашборд_ПБЕ»
'  (карточки, шкала, срезы, 4 графика, рейтинг, тепловая карта DSO) и лист сверки.
'
'  Запуск: Alt+F11 -> Insert -> Module -> вставить весь текст -> курсор в BuildPBE -> F5.
'  Работает 2-5 минут (читает папку «ДЗ и КЗ»). Можно запускать повторно — пересоздаёт своё.
'  Старые листы «Свод_ПБЕ» и «ПБЕ_данные (историч.)» НЕ трогает.
' =====================================================================
Option Explicit

Private LogText As String
Private wb As Workbook
Private fD0 As Object, fD1 As Object, fP1 As Object, fW As Object, fG As Object

Private Sub Chk(ByVal stepName As String)
    If Err.Number <> 0 Then
        LogText = LogText & "- " & stepName & ": " & Err.Description & vbLf
        Err.Clear
    End If
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    SheetExists = Not ws Is Nothing
    Err.Clear
End Function

%%QFUNCS%%

Private Sub AddQuery(ByVal nm As String, ByVal f As String)
    On Error Resume Next
    wb.Queries(nm).Delete
    Err.Clear
    On Error GoTo 0
    wb.Queries.Add nm, f
End Sub

Private Function ConnStr(ByVal nm As String) As String
    ConnStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & nm & ";Extended Properties="""""
End Function

Private Sub LoadToModel(ByVal nm As String)
    Dim c As WorkbookConnection
    Set c = wb.Connections.Add2("Запрос — " & nm, "Соединение с запросом " & nm, ConnStr(nm), _
        """" & nm & """", 6, True, False)
    c.Refresh
End Sub

Private Sub AddM(ByVal nm As String, ByVal f As String, fmt As Object)
    Dim t As Object
    Set t = wb.Model.ModelTables("ДПБ_Факт")
    On Error Resume Next
    wb.Model.ModelMeasures.Add nm, t, f, fmt
    If Err.Number <> 0 Then
        Err.Clear
        wb.Model.ModelMeasures.Add nm, t, Replace(f, ";", ","), fmt
    End If
    If Err.Number <> 0 Then
        LogText = LogText & "- мера «" & nm & "»: " & Err.Description & vbLf
        Err.Clear
    End If
End Sub

Private Function NewPivot(ws As Worksheet, ByVal addr As String, ByVal nm As String) As PivotTable
    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlExternal, _
        SourceData:=wb.Connections("ThisWorkbookDataModel"), Version:=6)
    Set NewPivot = pc.CreatePivotTable(TableDestination:=ws.Range(addr), TableName:=nm)
    With NewPivot
        .HasAutoFormat = False
        .PreserveFormatting = True
        .DisplayNullString = True
        .NullString = ChrW(8211)
        .TableStyle2 = "PivotStyleLight16"
    End With
End Function

Private Sub AddMeasure(pt As PivotTable, ByVal m As String, ByVal cap As String, ByVal fmt As String)
    pt.CubeFields("[Measures].[" & m & "]").Orientation = xlDataField
    With pt.DataFields(pt.DataFields.Count)
        If cap <> "" Then .Caption = cap
        If fmt <> "" Then .NumberFormat = fmt
    End With
End Sub

Private Sub TopN(pt As PivotTable, ByVal fld As String, ByVal m As String, ByVal n As Long)
    With pt.PivotFields(fld)
        .ClearAllFilters
        .PivotFilters.Add2 Type:=xlTopCount, DataField:=pt.CubeFields("[Measures].[" & m & "]"), Value1:=n
    End With
End Sub

Private Sub SortDesc(pt As PivotTable, ByVal fld As String, ByVal m As String)
    pt.PivotFields(fld).AutoSort xlDescending, "[Measures].[" & m & "]"
End Sub

Private Function AddPivotChart(ws As Worksheet, pt As PivotTable, ByVal ct As Long, _
        ByVal addr As String, ByVal ttl As String) As Chart
    Dim r As Range, shp As Shape
    Set r = ws.Range(addr)
    Set shp = ws.Shapes.AddChart2(-1, ct, r.Left, r.Top, r.Width, r.Height)
    With shp.Chart
        .SetSourceData Source:=pt.TableRange1
        .ShowAllFieldButtons = False
        .HasTitle = True
        .ChartTitle.Text = ttl
        .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 11
        .ChartTitle.Format.TextFrame2.TextRange.Font.Bold = msoTrue
    End With
    Set AddPivotChart = shp.Chart
End Function

Private Sub Connect(ByVal sc As SlicerCache, ByVal pt As PivotTable)
    Dim p As PivotTable
    If sc Is Nothing Or pt Is Nothing Then Exit Sub
    For Each p In sc.PivotTables
        If p.Name = pt.Name And p.Parent.Name = pt.Parent.Name Then Exit Sub
    Next p
    sc.PivotTables.AddPivotTable pt
End Sub

Private Sub DeleteSlicerCache(ByVal nm As String)
    On Error Resume Next
    wb.SlicerCaches(nm).Delete
    Err.Clear
End Sub

Private Sub FontColorBetween(r As Range, ByVal lo As Double, ByVal hi As Double, ByVal clr As Long)
    Dim fc As FormatCondition
    Set fc = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=" & CStr(lo), Formula2:="=" & CStr(hi))
    fc.Font.Color = clr
    fc.ScopeType = xlFieldsScope
End Sub

Private Sub BarChartSetup(ch As Chart, ByVal clr As Long, ByVal lblFmt As String, ByVal axFmt As String)
    ch.Axes(xlCategory).ReversePlotOrder = True
    ch.HasLegend = False
    With ch.SeriesCollection(1)
        .Format.Fill.ForeColor.RGB = clr
        .HasDataLabels = True
        .DataLabels.NumberFormat = lblFmt
    End With
    ch.Axes(xlValue).TickLabels.NumberFormat = axFmt
End Sub

Public Sub BuildPBE()
    Dim wsS As Worksheet, ws As Worksheet, wsV As Worksheet
    Dim ptK As PivotTable, ptD As PivotTable, ptE As PivotTable, ptSk As PivotTable, ptDSO As PivotTable
    Dim ptR As PivotTable, ptH As PivotTable
    Dim scT As SlicerCache, scP As SlicerCache, tl As SlicerCache
    Dim ch As Chart, r As Range, i As Long, c As String, c2 As String, sep As String
    Dim PB As String, PBF As String, TG As String, DASH As String, KP As String, tblV As String
    Dim cols, labels, measures, fmts, allPts, p, qn

    Set wb = ThisWorkbook
    LogText = ""
    PB = "[ДПБ_ПБЕ].[ПБЕ]"
    PBF = PB & ".[ПБЕ]"
    TG = ChrW(8376)
    DASH = ChrW(8211)
    KP = "'ДПБ_Своды'!$A$3"
    sep = Application.International(xlDecimalSeparator)

    If Not SheetExists("ПБЕ_данные (историч.)") Then MsgBox "Не найден лист «ПБЕ_данные (историч.)»", vbCritical: Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---------- 1. убрать результат прошлого запуска ----------
    DeleteSlicerCache "Срез_ДПБ_Тип"
    DeleteSlicerCache "Срез_ДПБ_ПБЕ"
    DeleteSlicerCache "Шкала_ДПБ"
    If SheetExists("Дашборд_ПБЕ") Then wb.Worksheets("Дашборд_ПБЕ").Delete
    If SheetExists("ДПБ_Своды") Then wb.Worksheets("ДПБ_Своды").Delete
    If SheetExists("ДПБ_Сверка") Then wb.Worksheets("ДПБ_Сверка").Delete
    For i = wb.Connections.Count To 1 Step -1
        If InStr(wb.Connections(i).Name, "ДПБ_") > 0 Then wb.Connections(i).Delete
    Next i
    For Each qn In Array("ДПБ_Сверка", "ДПБ_Параметры", "ДПБ_ПБЕ", "ДПБ_Календарь", "ДПБ_Факт", "ДПБ_Файлы", "ДПБ_История")
        wb.Queries(qn).Delete
    Next qn
    wb.Names("ДПБ_ИсторияДиапазон").Delete
    Err.Clear
    Application.DisplayAlerts = True

    ' ---------- 2. имя, листы, параметры ----------
    wb.Names.Add Name:="ДПБ_ИсторияДиапазон", RefersTo:="='ПБЕ_данные (историч.)'!$B$6:$I$20000"
    Chk "именованный диапазон"
    Set wsS = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsS.Name = "ДПБ_Своды"
    Set wsV = wb.Worksheets.Add(After:=wsS)
    wsV.Name = "ДПБ_Сверка"
    If SheetExists("Карта_DIO") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Карта_DIO"))
    Else
        Set ws = wb.Worksheets.Add(Before:=wsS)
    End If
    ws.Name = "Дашборд_ПБЕ"
    wsS.Range("R3").Value = "Цель_DSO"
    wsS.Range("S3").Value = "Порог_доли"
    wsS.Range("R4").Formula = "=Параметры!C21"
    wsS.Range("S4").Value = 0.001
    wsS.ListObjects.Add(xlSrcRange, wsS.Range("R3:S4"), , xlYes).Name = "tblДПБ_Параметры"
    Chk "листы и параметры"

    ' ---------- 3. запросы Power Query ----------
    AddQuery "ДПБ_История", QM0()
    AddQuery "ДПБ_Файлы", QM1()
    AddQuery "ДПБ_Факт", QM2()
    AddQuery "ДПБ_Календарь", QM3()
    AddQuery "ДПБ_ПБЕ", QM4()
    AddQuery "ДПБ_Параметры", QM5()
    AddQuery "ДПБ_Сверка", QM6()
    Chk "создание запросов"
    If LogText <> "" Then GoTo Finish

    ' ---------- 4. загрузка в модель данных ----------
    LoadToModel "ДПБ_Факт":      Chk "загрузка ДПБ_Факт"
    LoadToModel "ДПБ_Календарь": Chk "загрузка ДПБ_Календарь"
    LoadToModel "ДПБ_ПБЕ":       Chk "загрузка ДПБ_ПБЕ"
    LoadToModel "ДПБ_Параметры": Chk "загрузка ДПБ_Параметры"
    If LogText <> "" Then GoTo Finish

    ' ---------- 5. лист сверки ----------
    With wsV.ListObjects.Add(SourceType:=0, Source:=ConnStr("ДПБ_Сверка"), Destination:=wsV.Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [ДПБ_Сверка]")
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = True
        .ListObject.DisplayName = "ДПБ_Сверка"
        .Refresh BackgroundQuery:=False
    End With
    Chk "лист сверки"
    tblV = wsV.ListObjects(1).Name
    Err.Clear

    ' ---------- 6. связи ----------
    With wb.Model
        .ModelRelationships.Add .ModelTables("ДПБ_Факт").ModelTableColumns("Период"), _
                                .ModelTables("ДПБ_Календарь").ModelTableColumns("Дата")
        .ModelRelationships.Add .ModelTables("ДПБ_Факт").ModelTableColumns("ПБЕ"), _
                                .ModelTables("ДПБ_ПБЕ").ModelTableColumns("ПБЕ")
    End With
    Chk "связи в модели"

    ' ---------- 7. меры ----------
    Set fD0 = wb.Model.ModelFormatDecimalNumber(True, 0)
    Set fD1 = wb.Model.ModelFormatDecimalNumber(True, 1)
    Set fP1 = wb.Model.ModelFormatPercentageNumber(False, 1)
    Set fW = wb.Model.ModelFormatWholeNumber(True)
    Set fG = wb.Model.ModelFormatGeneral
    Chk "форматы мер"
%%MEASURES%%
    If LogText <> "" Then GoTo Finish

    ' ---------- 8. служебные сводные ----------
    Set ptK = NewPivot(wsS, "A3", "П2_KPI")
    For Each p In Array("Период ПБЕ текст", "Оборот Дт млн", "Прирост оборота гг", "ДЗ на конец млн", "DSO", _
                        "Дельта DSO", "Инкассация", "Авансы на конец млн", "Излишек ДЗ млн", "Цель DSO")
        ptK.CubeFields("[Measures].[" & p & "]").Orientation = xlDataField
    Next p
    Chk "П2_KPI"
    Set ptD = NewPivot(wsS, "A10", "П2_Динамика")
    ptD.CubeFields("[ДПБ_Календарь].[Месяц]").Orientation = xlRowField
    AddMeasure ptD, "Оборот Дт млн", "Оборот, млн " & TG, "#,##0.0"
    AddMeasure ptD, "DSO", "DSO, дн.", "0.0"
    Chk "П2_Динамика"
    Set ptE = NewPivot(wsS, "E10", "П2_Излишек")
    ptE.CubeFields(PB).Orientation = xlRowField
    AddMeasure ptE, "Излишек ДЗ млн", "Излишек ДЗ, млн " & TG, "#,##0.0"
    TopN ptE, PBF, "Излишек ДЗ млн", 15
    SortDesc ptE, PBF, "Излишек ДЗ млн"
    Chk "П2_Излишек"
    Set ptSk = NewPivot(wsS, "I10", "П2_Перекос")
    ptSk.CubeFields(PB).Orientation = xlRowField
    AddMeasure ptSk, "Перекос ДЗ", "Перекос (ДЗ - оборот)", "+0.0%;-0.0%;0.0%"
    TopN ptSk, PBF, "Перекос ДЗ", 15
    SortDesc ptSk, PBF, "Перекос ДЗ"
    Chk "П2_Перекос"
    Set ptDSO = NewPivot(wsS, "M10", "П2_DSO")
    ptDSO.CubeFields(PB).Orientation = xlRowField
    AddMeasure ptDSO, "DSO", "DSO за период, дн.", "0.0"
    SortDesc ptDSO, PBF, "DSO"
    Chk "П2_DSO"

    ' ================= ЛИСТ ДАШБОРДА =================
    With ws
        .Columns("A").ColumnWidth = 2
        .Columns("B:D").ColumnWidth = 13
        .Columns("E").ColumnWidth = 2
        .Columns("F").ColumnWidth = 30
        .Columns("G:R").ColumnWidth = 11
        .Rows("5:8").RowHeight = 26
        .Range("B2").Value = "ПБЕ — ДЕБИТОРСКАЯ ЗАДОЛЖЕННОСТЬ ПОКУПАТЕЛЕЙ И СКОРОСТЬ ОПЛАТ"
        .Range("B2").Font.Size = 20
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(31, 56, 100)
        .Range("B3").Formula = "=""Период: ""&IFERROR(GETPIVOTDATA(""[Measures].[Период ПБЕ текст]""," & KP & "),""" & DASH & """)"
        .Range("B3").Font.Color = RGB(89, 89, 89)
        If tblV <> "" Then
            .Range("M3").Formula = "=IF(COUNTIF(" & tblV & "[Статус],""РАСХОЖДЕНИЕ"")+COUNTIF(" & tblV & _
                "[Статус],""история: файлы НЕ сходятся"")=0,""" & ChrW(10004) & " Данные сверены"",""" & ChrW(9888) & _
                " Есть расхождения — см. лист ДПБ_Сверка"")"
            .Range("M3").Font.Bold = True
            .Range("M3").Font.Color = RGB(46, 125, 50)
        End If
        .Range("B4").Value = "DSO = средняя ДЗ покупателей " & ChrW(215) & " дни периода " & ChrW(247) & _
            " отгрузки (оборот по дебету счетов 121х/1260, с НДС). «Без метки ПБЕ» — розница и маркетплейсы."
        .Range("B4").Font.Size = 9
        .Range("B4").Font.Italic = True
        .Range("B4").Font.Color = RGB(120, 120, 120)
    End With
    Chk "заголовок"

    cols = Array("F", "H", "J", "L", "N", "P")
    labels = Array("Отгрузки (оборот Дт), млн " & TG, "ДЗ покупателей на конец, млн " & TG, "DSO, дней", _
                   "Инкассация: оплаты " & ChrW(247) & " отгрузки", "Авансы покупателей, млн " & TG, "Излишек ДЗ сверх цели DSO, млн " & TG)
    measures = Array("Оборот Дт млн", "ДЗ на конец млн", "DSO", "Инкассация", "Авансы на конец млн", "Излишек ДЗ млн")
    fmts = Array("#,##0.0", "#,##0.0", "0.0", "0.0%", "#,##0.0", "#,##0.0")
    For i = 0 To 5
        c = cols(i): c2 = Chr(Asc(c) + 1)
        With ws.Range(c & "10:" & c2 & "13")
            .Interior.Color = RGB(242, 244, 248)
            .BorderAround xlContinuous, xlThin, , RGB(208, 215, 226)
        End With
        ws.Range(c & "10").Value = labels(i)
        ws.Range(c & "10").Font.Size = 9
        ws.Range(c & "10").Font.Color = RGB(89, 89, 89)
        With ws.Range(c & "11:" & c2 & "12")
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 22
            .Font.Bold = True
            .Font.Color = RGB(31, 56, 100)
            .NumberFormat = fmts(i)
        End With
        ws.Range(c & "11").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[" & measures(i) & "]""," & KP & "),""" & DASH & """)"
        ws.Range(c & "13:" & c2 & "13").Merge
        ws.Range(c & "13").HorizontalAlignment = xlCenter
        ws.Range(c & "13").Font.Size = 9
    Next i
    ws.Range("F13").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[Прирост оборота гг]""," & KP & "),"""")"
    ws.Range("F13").NumberFormat = "[Color10]+0.0%"" к прошлому году"";[Red]-0.0%"" к прошлому году"";0.0%"" к прошлому году"""
    ws.Range("J13").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[Дельта DSO]""," & KP & "),"""")"
    ws.Range("J13").NumberFormat = "[Red]+0.0"" дн. к пред. периоду"";[Color10]-0.0"" дн. к пред. периоду"";0.0"" дн. к пред. периоду"""
    ws.Range("P13").Formula = "=""цель DSO ""&'ДПБ_Своды'!$R$4&"" дн."""
    Chk "карточки"

    ' ---------- рейтинг ----------
    ws.Range("F50").Value = "РЕЙТИНГ ПБЕ ПО ОБОРОТУ ЗА ВЫБРАННЫЙ ПЕРИОД"
    ws.Range("F50").Font.Size = 12
    ws.Range("F50").Font.Bold = True
    ws.Range("F50").Font.Color = RGB(31, 56, 100)
    Set ptR = NewPivot(ws, "F51", "П2_Рейтинг")
    ptR.CubeFields(PB).Orientation = xlRowField
    AddMeasure ptR, "Ранг ПБЕ", "№", "0"
    AddMeasure ptR, "Оборот Дт млн", "Отгрузки, млн " & TG, "#,##0.0"
    AddMeasure ptR, "Доля оборота", "Доля в отгрузках", "0.0%"
    AddMeasure ptR, "Прирост оборота", "Прирост к пред. периоду", "+0.0%;-0.0%;0.0%"
    AddMeasure ptR, "Средняя ДЗ млн", "Ср. ДЗ, млн " & TG, "#,##0.0"
    AddMeasure ptR, "Доля ДЗ", "Доля в ДЗ", "0.0%"
    AddMeasure ptR, "DSO", "DSO, дн.", "0.0"
    AddMeasure ptR, "Дельта DSO", ChrW(916) & " DSO, дн.", "+0.0;-0.0;0.0"
    AddMeasure ptR, "Инкассация", "Инкассация, %", "0.0%"
    AddMeasure ptR, "Авансы на конец млн", "Авансы, млн " & TG, "#,##0.0"
    AddMeasure ptR, "Излишек ДЗ млн", "Излишек ДЗ, млн " & TG, "#,##0.0"
    AddMeasure ptR, "Перекос ДЗ", "Перекос (ДЗ - отгрузки)", "+0.0%;-0.0%;0.0%"
    ptR.RowAxisLayout xlTabularRow
    SortDesc ptR, PBF, "Оборот Дт млн"
    With ptR.TableRange1.Rows(1)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 42
    End With
    Chk "рейтинг"
    Set r = ptR.DataFields(2).DataRange
    With r.FormatConditions.AddDatabar
        .BarColor.Color = RGB(46, 117, 182)
        .BarFillType = xlDataBarFillSolid
        .ScopeType = xlFieldsScope
    End With
    Set r = ptR.DataFields(7).DataRange
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(99, 190, 123)
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
        .ScopeType = xlFieldsScope
    End With
    Set r = ptR.DataFields(8).DataRange
    FontColorBetween r, 0.05, 100000, RGB(198, 40, 40)
    FontColorBetween r, -100000, -0.05, RGB(46, 125, 50)
    Set r = ptR.DataFields(11).DataRange
    With r.FormatConditions.AddDatabar
        .BarColor.Color = RGB(229, 115, 115)
        .BarFillType = xlDataBarFillSolid
        .ScopeType = xlFieldsScope
    End With
    Set r = ptR.DataFields(12).DataRange
    With r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=" & CStr(0.001), Formula2:="=10")
        .Interior.Color = RGB(255, 228, 228)
        .ScopeType = xlFieldsScope
    End With
    Chk "подсветка рейтинга"

    ' ---------- тепловая карта DSO ----------
    ws.Range("F89").Value = "ТЕПЛОВАЯ КАРТА DSO ПО МЕСЯЦАМ: зелёный — платят быстро, красный — медленно"
    ws.Range("F89").Font.Size = 12
    ws.Range("F89").Font.Bold = True
    ws.Range("F89").Font.Color = RGB(31, 56, 100)
    Set ptH = NewPivot(ws, "F90", "П2_Карта")
    ptH.CubeFields(PB).Orientation = xlRowField
    ptH.CubeFields("[ДПБ_Календарь].[Месяц]").Orientation = xlColumnField
    AddMeasure ptH, "DSO", "DSO, дн.", "0"
    ptH.RowAxisLayout xlTabularRow
    Set r = ptH.DataFields(1).DataRange
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(99, 190, 123)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).Type = xlConditionValuePercentile
        .ColorScaleCriteria(3).Value = 95
        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
        .ScopeType = xlFieldsScope
    End With
    Chk "тепловая карта"

    ' ---------- графики ----------
    ws.Activate
    ws.Range("A1").Select
    Set ch = AddPivotChart(ws, ptD, xlColumnClustered, "F15:K31", _
        "Отгрузки (млн " & TG & ") и DSO (дней) по месяцам — выбранные ПБЕ, вся история")
    With ch.SeriesCollection(2)
        .ChartType = xlLine
        .AxisGroup = xlSecondary
        .Format.Line.ForeColor.RGB = RGB(237, 125, 49)
        .Format.Line.Weight = 2.25
    End With
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(46, 117, 182)
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    Chk "график динамики"
    Set ch = AddPivotChart(ws, ptE, xlBarClustered, "L15:Q31", "Топ-15: излишек ДЗ сверх цели DSO, млн " & TG)
    BarChartSetup ch, RGB(198, 40, 40), "0" & sep & "0", "0"
    Chk "график излишка"
    Set ch = AddPivotChart(ws, ptSk, xlBarClustered, "F32:K48", "Топ-15 по перекосу: доля в ДЗ выше доли в отгрузках")
    BarChartSetup ch, RGB(237, 125, 49), "+0" & sep & "0%", "0%"
    Chk "график перекоса"
    Set ch = AddPivotChart(ws, ptDSO, xlBarClustered, "L32:Q48", "DSO по ПБЕ за выбранный период, дней")
    BarChartSetup ch, RGB(46, 117, 182), "0", "0"
    Chk "график DSO"

    ' ---------- срезы и шкала ----------
    allPts = Array(ptK, ptD, ptE, ptSk, ptDSO, ptR, ptH)
    Set scT = wb.SlicerCaches.Add2(ptK, "[ДПБ_ПБЕ].[Тип]", "Срез_ДПБ_Тип")
    Set scP = wb.SlicerCaches.Add2(ptK, PB, "Срез_ДПБ_ПБЕ")
    scP.SlicerCacheLevels(1).CrossFilterType = xlSlicerCrossFilterHideButtonsWithNoData
    Chk "срезы"
    scT.Slicers.Add ws, "[ДПБ_ПБЕ].[Тип].[Тип]", "ДПБ_Тип_Д", "Тип", _
        ws.Range("B5").Top, ws.Range("B5").Left, ws.Range("B5:D5").Width, 80
    scP.Slicers.Add ws, PBF, "ДПБ_ПБЕ_Д", "ПБЕ", _
        ws.Range("B5").Top + 88, ws.Range("B5").Left, ws.Range("B5:D5").Width, ws.Range("B14:B48").Height + 40
    Chk "срезы на листе"
    For Each p In allPts
        Connect scT, p
        Connect scP, p
    Next p
    Chk "подключение срезов"
    Set tl = wb.SlicerCaches.Add2(ptK, "[ДПБ_Календарь].[Дата]", "Шкала_ДПБ", xlTimeline)
    tl.Slicers.Add ws, , "ДПБ_Шкала_Д", "Период", ws.Range("F5").Top, ws.Range("F5").Left, ws.Range("F5:Q5").Width, ws.Range("F5:F8").Height
    Chk "временная шкала"
    For Each p In allPts
        If p.Name <> "П2_Динамика" Then Connect tl, p
    Next p
    tl.PivotTables.RemovePivotTable ptD
    Err.Clear
    Chk "подключение шкалы"

Finish:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If SheetExists("ДПБ_Своды") And SheetExists("Дашборд_ПБЕ") Then
        wb.Worksheets("ДПБ_Своды").Visible = xlSheetHidden
        wb.Worksheets("Дашборд_ПБЕ").Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.Zoom = 80
        wb.Worksheets("Дашборд_ПБЕ").Range("A1").Select
    End If
    Err.Clear
    On Error GoTo 0
    If LogText = "" Then
        MsgBox "Готово! Дашборд ПБЕ собран без ошибок." & vbLf & "Проверьте лист ДПБ_Сверка.", vbInformation
    Else
        MsgBox "Сборка остановлена или прошла не полностью. Пришлите этот текст:" & vbLf & vbLf & LogText, vbExclamation
    End If
End Sub
