' =====================================================================
'  Сборка листов «Дашборд_представительства» и «Карта_DIO»
'
'  Что нужно ДО запуска (шаги 1–7 инструкции):
'    - лист ДП_Своды со сводными П_KPI, П_Динамика, П_Излишек, П_Перекос, П_ABC;
'    - лист ДП_Сверка с таблицей сверки;
'    - таблица tblДП_Параметры на листе ДП_Своды (R3:S4).
'
'  Как запустить: Alt+F11 -> Insert -> Module -> вставить весь этот текст ->
'  курсор внутрь процедуры BuildRepDashboard -> F5.
'  Макрос можно запускать повторно: он пересоздаёт свои листы и срезы.
'  После запуска файл сохраняйте как обычно (.xlsx): на вопрос о макросах
'  ответьте «Да» — макрос одноразовый, в файле он не нужен.
' =====================================================================
Option Explicit

Private LogText As String
Private wb As Workbook

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
        .PivotFilters.Add2 Type:=xlTopCount, _
            DataField:=pt.CubeFields("[Measures].[" & m & "]"), Value1:=n
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
    ' формулы условного форматирования Excel читает в местном формате (десятичная запятая) — CStr даёт именно его
    Set fc = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=" & CStr(lo), Formula2:="=" & CStr(hi))
    fc.Font.Color = clr
    fc.ScopeType = xlFieldsScope
End Sub

Public Sub BuildRepDashboard()
    Dim wsS As Worksheet, ws As Worksheet, wsM As Worksheet
    Dim ptKPI As PivotTable, ptDyn As PivotTable, ptExc As PivotTable, ptSkew As PivotTable, ptABC As PivotTable
    Dim ptRank As PivotTable, ptTop As PivotTable, ptMap As PivotTable
    Dim scSt As SlicerCache, scABC As SlicerCache, scRep As SlicerCache, tl As SlicerCache, sc As SlicerCache
    Dim ch As Chart, r As Range, i As Long, c As String, c2 As String
    Dim REP As String, REPF As String, TG As String, DASH As String, tblSv As String
    Dim cols, labels, measures, fmts, allPts, p

    Set wb = ThisWorkbook
    LogText = ""
    REP = "[ДП_Представительства].[Представительство]"
    REPF = REP & ".[Представительство]"
    TG = ChrW(8376)          ' знак тенге
    DASH = ChrW(8211)        ' тире

    ' ---------- проверки ----------
    If Not SheetExists("ДП_Своды") Then MsgBox "Не найден лист ДП_Своды", vbCritical: Exit Sub
    Set wsS = wb.Worksheets("ДП_Своды")
    On Error Resume Next
    Set ptKPI = wsS.PivotTables("П_KPI")
    Set ptDyn = wsS.PivotTables("П_Динамика")
    Set ptExc = wsS.PivotTables("П_Излишек")
    Set ptSkew = wsS.PivotTables("П_Перекос")
    Set ptABC = wsS.PivotTables("П_ABC")
    Err.Clear
    If ptKPI Is Nothing Or ptDyn Is Nothing Or ptExc Is Nothing Or ptSkew Is Nothing Or ptABC Is Nothing Then
        On Error GoTo 0
        MsgBox "На листе ДП_Своды не найдены все 5 сводных (П_KPI, П_Динамика, П_Излишек, П_Перекос, П_ABC)." & vbLf & _
               "Проверьте их имена: правый клик по сводной -> Параметры сводной таблицы -> Имя.", vbCritical
        Exit Sub
    End If
    tblSv = ""
    tblSv = wb.Worksheets("ДП_Сверка").ListObjects(1).Name
    Err.Clear

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---------- убрать то, что мог создать прошлый запуск ----------
    DeleteSlicerCache "Срез_ДП_Статус"
    DeleteSlicerCache "Срез_ДП_ABC"
    DeleteSlicerCache "Срез_ДП_Представительство"
    If SheetExists("Дашборд_предст_СТАРЫЙ") Then
        If SheetExists("Дашборд_представительства") Then wb.Worksheets("Дашборд_представительства").Delete
    ElseIf SheetExists("Дашборд_представительства") Then
        wb.Worksheets("Дашборд_представительства").Name = "Дашборд_предст_СТАРЫЙ"
    End If
    If SheetExists("Карта_DIO") Then wb.Worksheets("Карта_DIO").Delete
    Chk "подготовка листов"
    Application.DisplayAlerts = True

    ' ---------- новые листы ----------
    If SheetExists("Дашборд") Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Дашборд"))
    Else
        Set ws = wb.Worksheets.Add(After:=wsS)
    End If
    ws.Name = "Дашборд_представительства"
    Set wsM = wb.Worksheets.Add(After:=ws)
    wsM.Name = "Карта_DIO"
    Chk "создание листов"

    ' ================= ЛИСТ ДАШБОРДА =================
    With ws
        .Columns("A").ColumnWidth = 2
        .Columns("B:D").ColumnWidth = 13
        .Columns("E").ColumnWidth = 2
        .Columns("F").ColumnWidth = 30
        .Columns("G:Q").ColumnWidth = 11
        .Rows("5:8").RowHeight = 26

        .Range("B2").Value = "ПРЕДСТАВИТЕЛЬСТВА — ПРОДАЖИ, ЗАПАСЫ И ОБОРАЧИВАЕМОСТЬ"
        .Range("B2").Font.Size = 20
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(31, 56, 100)

        .Range("B3").Formula = "=""Период: ""&IFERROR(GETPIVOTDATA(""[Measures].[Период текст]"",'ДП_Своды'!$A$3),""" & DASH & """)"
        .Range("B3").Font.Size = 11
        .Range("B3").Font.Color = RGB(89, 89, 89)

        If tblSv <> "" Then
            .Range("M3").Formula = "=IF(COUNTIF(" & tblSv & "[Статус],""РАСХОЖДЕНИЕ"")=0,""" & ChrW(10004) & _
                " Данные сверены с общей моделью"",""" & ChrW(9888) & " Есть расхождения — см. лист ДП_Сверка"")"
            .Range("M3").Font.Bold = True
            .Range("M3").Font.Color = RGB(46, 125, 50)
        End If
    End With
    Chk "заголовок"

    ' ---------- карточки KPI ----------
    cols = Array("F", "H", "J", "L", "N", "P")
    labels = Array("Выручка, млн " & TG, "Валовая маржа", "Средние запасы, млн " & TG, _
                   "DIO, дней", "Излишек запасов, млн " & TG, "Активных представительств")
    measures = Array("Выручка млн", "Маржа доля", "Средние запасы млн", "DIO", "Излишек запасов млн", "Активных представительств")
    fmts = Array("#,##0.0", "0.0%", "#,##0.0", "0.0", "#,##0.0", "0")
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
        ws.Range(c & "11").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[" & measures(i) & "]"",'ДП_Своды'!$A$3),""" & DASH & """)"
        ws.Range(c & "13").Font.Size = 9
    Next i
    ws.Range("F13").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[Прирост выручки гг]"",'ДП_Своды'!$A$3),"""")"
    ws.Range("F13").NumberFormat = "[Color10]+0.0%"" к прошлому году"";[Red]-0.0%"" к прошлому году"";0.0%"" к прошлому году"""
    ws.Range("L13").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[Дельта DIO]"",'ДП_Своды'!$A$3),"""")"
    ws.Range("L13").NumberFormat = "[Red]+0.0"" дн. к пред. периоду"";[Color10]-0.0"" дн. к пред. периоду"";0.0"" дн. к пред. периоду"""
    ws.Range("N13").Formula = "=""сверх цели DIO ""&'ДП_Своды'!$R$4&"" дн."""
    Chk "карточки KPI"

    ' ---------- рейтинг (внизу листа) ----------
    ws.Range("F50").Value = "РЕЙТИНГ ПРЕДСТАВИТЕЛЬСТВ — ТОП-25 ПО ВЫРУЧКЕ ЗА ВЫБРАННЫЙ ПЕРИОД"
    ws.Range("F50").Font.Size = 12
    ws.Range("F50").Font.Bold = True
    ws.Range("F50").Font.Color = RGB(31, 56, 100)
    Set ptRank = NewPivot(ws, "F51", "П_Рейтинг")
    Chk "создание П_Рейтинг"
    ptRank.CubeFields(REP).Orientation = xlRowField
    AddMeasure ptRank, "Ранг по выручке", "№", "0"
    AddMeasure ptRank, "Выручка млн", "Выручка, млн " & TG, "#,##0.0"
    AddMeasure ptRank, "Доля выручки", "Доля в выручке", "0.0%"
    AddMeasure ptRank, "Прирост выручки", "Прирост к пред. периоду", "+0.0%;-0.0%;0.0%"
    AddMeasure ptRank, "Маржа доля", "Маржа", "0.0%"
    AddMeasure ptRank, "Средние запасы млн", "Ср. запасы, млн " & TG, "#,##0.0"
    AddMeasure ptRank, "DIO", "DIO, дн.", "0.0"
    AddMeasure ptRank, "Дельта DIO", ChrW(916) & " DIO, дн.", "+0.0;-0.0;0.0"
    AddMeasure ptRank, "Доля запасов", "Доля в запасах", "0.0%"
    AddMeasure ptRank, "Перекос", "Перекос (запасы - выручка)", "+0.0%;-0.0%;0.0%"
    AddMeasure ptRank, "Излишек запасов млн", "Излишек, млн " & TG, "#,##0.0"
    ptRank.RowAxisLayout xlTabularRow
    Chk "поля П_Рейтинг"
    TopN ptRank, REPF, "Выручка млн", 25
    Chk "топ-25 в рейтинге"
    SortDesc ptRank, REPF, "Выручка млн"
    Chk "сортировка рейтинга"

    ' подсветка рейтинга
    Set r = ptRank.DataFields(2).DataRange
    With r.FormatConditions.AddDatabar
        .BarColor.Color = RGB(46, 117, 182)
        .BarFillType = xlDataBarFillSolid
        .ScopeType = xlFieldsScope
    End With
    Chk "гистограмма выручки"
    Set r = ptRank.DataFields(7).DataRange
    With r.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(99, 190, 123)
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
        .ScopeType = xlFieldsScope
    End With
    Chk "шкала DIO"
    Set r = ptRank.DataFields(8).DataRange
    FontColorBetween r, 0.05, 100000, RGB(198, 40, 40)
    FontColorBetween r, -100000, -0.05, RGB(46, 125, 50)
    Chk "цвет изменения DIO"
    Set r = ptRank.DataFields(10).DataRange
    With r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=" & CStr(0.001), Formula2:="=10")
        .Interior.Color = RGB(255, 228, 228)
        .ScopeType = xlFieldsScope
    End With
    Chk "подсветка перекоса"
    Set r = ptRank.DataFields(11).DataRange
    With r.FormatConditions.AddDatabar
        .BarColor.Color = RGB(229, 115, 115)
        .BarFillType = xlDataBarFillSolid
        .ScopeType = xlFieldsScope
    End With
    Chk "гистограмма излишка"

    ' ---------- графики ----------
    ws.Activate
    ws.Range("A1").Select
    Set ch = AddPivotChart(ws, ptDyn, xlColumnClustered, "F15:K31", _
        "Выручка (млн " & TG & ") и DIO (дней) по месяцам — выбранные представительства, вся история")
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

    Set ch = AddPivotChart(ws, ptExc, xlBarClustered, "L15:Q31", _
        "Топ-15: излишек запасов сверх цели DIO, млн " & TG)
    ch.Axes(xlCategory).ReversePlotOrder = True
    ch.HasLegend = False
    With ch.SeriesCollection(1)
        .Format.Fill.ForeColor.RGB = RGB(198, 40, 40)
        .HasDataLabels = True
        .DataLabels.NumberFormat = "#,##0"
    End With
    Chk "график излишка"

    Set ch = AddPivotChart(ws, ptSkew, xlBarClustered, "F32:K48", _
        "Топ-15 по перекосу: доля в запасах выше доли в выручке")
    ch.Axes(xlCategory).ReversePlotOrder = True
    ch.HasLegend = False
    With ch.SeriesCollection(1)
        .Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        .HasDataLabels = True
        .DataLabels.NumberFormat = "+0.0%"
    End With
    Chk "график перекоса"

    Set ch = AddPivotChart(ws, ptABC, xlDoughnut, "L32:Q48", "Структура выручки по классам ABC")
    With ch.SeriesCollection(1)
        .HasDataLabels = True
        .DataLabels.ShowValue = False
        .DataLabels.ShowPercentage = True
        .DataLabels.ShowCategoryName = False
    End With
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionRight
    Chk "график ABC"

    ' ================= ЛИСТ КАРТА_DIO =================
    With wsM
        .Columns("A").ColumnWidth = 2
        .Columns("B:D").ColumnWidth = 13
        .Columns("E").ColumnWidth = 2
        .Columns("F").ColumnWidth = 30
        .Columns("G:BZ").ColumnWidth = 9
        .Rows("3:7").RowHeight = 21
        .Range("B2").Value = "КАРТА DIO — ГДЕ ЗАПАСЫ ОБОРАЧИВАЮТСЯ МЕДЛЕННЕЕ"
        .Range("B2").Font.Size = 18
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Color = RGB(31, 56, 100)
        .Range("F9").Value = "Топ-10: где DIO вырос сильнее всего к предыдущему периоду (представительства с долей выручки от 0,1%)"
        .Range("F9").Font.Bold = True
        .Range("F25").Value = "Тепловая карта DIO по месяцам: зелёный — запасы оборачиваются быстро, красный — медленно"
        .Range("F25").Font.Bold = True
    End With
    Set ptTop = NewPivot(wsM, "F10", "П_ТопDIO")
    Chk "создание П_ТопDIO"
    ptTop.CubeFields(REP).Orientation = xlRowField
    AddMeasure ptTop, "DIO пред период", "DIO было", "0.0"
    AddMeasure ptTop, "DIO", "DIO стало", "0.0"
    AddMeasure ptTop, "Дельта DIO значимая", "Рост DIO, дн.", "+0.0;-0.0;0.0"
    ptTop.RowAxisLayout xlTabularRow
    Chk "поля П_ТопDIO"
    TopN ptTop, REPF, "Дельта DIO значимая", 10
    Chk "топ-10 роста DIO"
    SortDesc ptTop, REPF, "Дельта DIO значимая"
    Chk "сортировка топ-10"

    Set ptMap = NewPivot(wsM, "F26", "П_Карта")
    Chk "создание П_Карта"
    ptMap.CubeFields(REP).Orientation = xlRowField
    ptMap.CubeFields("[ДП_Календарь].[Месяц]").Orientation = xlColumnField
    AddMeasure ptMap, "DIO", "DIO, дн.", "0"
    ptMap.RowAxisLayout xlTabularRow
    Chk "поля П_Карта"
    Set r = ptMap.DataFields(1).DataRange
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

    ' ================= СРЕЗЫ И ВРЕМЕННАЯ ШКАЛА =================
    allPts = Array(ptKPI, ptDyn, ptExc, ptSkew, ptABC, ptRank, ptTop, ptMap)

    Set scSt = wb.SlicerCaches.Add2(ptKPI, "[ДП_Представительства].[Статус]", "Срез_ДП_Статус")
    Chk "срез Статус"
    Set scABC = wb.SlicerCaches.Add2(ptKPI, "[ДП_Представительства].[Класс ABC]", "Срез_ДП_ABC")
    Chk "срез Класс ABC"
    Set scRep = wb.SlicerCaches.Add2(ptKPI, REP, "Срез_ДП_Представительство")
    Chk "срез Представительство"
    scRep.SlicerCacheLevels(1).CrossFilterType = xlSlicerCrossFilterHideButtonsWithNoData
    scABC.SlicerCacheLevels(1).CrossFilterType = xlSlicerCrossFilterHideButtonsWithNoData
    Chk "скрытие пустых элементов"

    ' визуальные срезы на дашборде (слева, B:D) и такие же на Карта_DIO — они синхронны
    With ws
        scSt.Slicers.Add ws, "[ДП_Представительства].[Статус].[Статус]", "ДП_Статус_Д", "Статус", _
            .Range("B5").Top, .Range("B5").Left, .Range("B5:D5").Width, 100
        scABC.Slicers.Add ws, "[ДП_Представительства].[Класс ABC].[Класс ABC]", "ДП_ABC_Д", "Класс ABC", _
            .Range("B5").Top + 108, .Range("B5").Left, .Range("B5:D5").Width, 118
        scRep.Slicers.Add ws, REPF, "ДП_Предст_Д", "Представительство", _
            .Range("B5").Top + 234, .Range("B5").Left, .Range("B5:D5").Width, .Range("B20:B48").Height + 60
    End With
    Chk "срезы на дашборде"
    With wsM
        scSt.Slicers.Add wsM, "[ДП_Представительства].[Статус].[Статус]", "ДП_Статус_К", "Статус", _
            .Range("B9").Top, .Range("B9").Left, .Range("B9:D9").Width, 100
        scABC.Slicers.Add wsM, "[ДП_Представительства].[Класс ABC].[Класс ABC]", "ДП_ABC_К", "Класс ABC", _
            .Range("B9").Top + 108, .Range("B9").Left, .Range("B9:D9").Width, 118
        scRep.Slicers.Add wsM, REPF, "ДП_Предст_К", "Представительство", _
            .Range("B9").Top + 234, .Range("B9").Left, .Range("B9:D9").Width, 420
    End With
    Chk "срезы на Карта_DIO"

    ' подключения срезов
    For Each p In allPts
        Connect scSt, p
        Connect scRep, p
        If p.Name <> "П_ABC" Then Connect scABC, p
    Next p
    Chk "подключение срезов"

    ' временная шкала: берём существующую (созданную на шаге 7.3) или создаём
    Set tl = Nothing
    For Each sc In wb.SlicerCaches
        If sc.SlicerCacheType = xlTimeline Then Set tl = sc: Exit For
    Next sc
    If tl Is Nothing Then
        Set tl = wb.SlicerCaches.Add2(ptKPI, "[ДП_Календарь].[Дата]", "Шкала_ДП", xlTimeline)
    End If
    Chk "поиск временной шкалы"
    tl.Slicers.Add ws, , "ДП_Шкала_Д", "Период", ws.Range("F5").Top, ws.Range("F5").Left, ws.Range("F5:Q5").Width, ws.Range("F5:F8").Height
    Chk "шкала на дашборде"
    tl.Slicers.Add wsM, , "ДП_Шкала_К", "Период", wsM.Range("F3").Top, wsM.Range("F3").Left, wsM.Range("F3:Q3").Width, wsM.Range("F3:F7").Height
    Chk "шкала на Карта_DIO"
    For Each p In allPts
        If p.Name <> "П_Динамика" Then Connect tl, p
    Next p
    tl.PivotTables.RemovePivotTable ptDyn   ' график динамики всегда показывает всю историю
    Err.Clear
    Chk "подключение шкалы"

    ' ---------- финал ----------
    Application.ScreenUpdating = True
    wsM.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    wsM.Range("A1").Select
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 80
    ws.Range("A1").Select
    On Error GoTo 0

    If LogText = "" Then
        MsgBox "Готово! Дашборд и Карта_DIO собраны без ошибок.", vbInformation
    Else
        MsgBox "Готово, но эти пункты не получились — их нужно сделать вручную или прислать мне:" & vbLf & vbLf & LogText, vbExclamation
    End If
End Sub

' =====================================================================
'  Косметические исправления после первого запуска (запускать отдельно)
' =====================================================================
Public Sub FixDashboard()
    Dim ws As Worksheet, co As ChartObject, t As String, sep As String, pt As PivotTable
    Set ws = ThisWorkbook.Worksheets("Дашборд_представительства")
    sep = Application.International(xlDecimalSeparator)
    On Error Resume Next
    ' строки-пояснения под карточками: объединить на 2 столбца, чтобы текст помещался
    ws.Range("F13:G13").Merge
    ws.Range("L13:M13").Merge
    ws.Range("N13:O13").Merge
    ws.Range("F13,L13,N13").HorizontalAlignment = xlCenter
    ' подписи и оси графиков
    For Each co In ws.ChartObjects
        t = co.Chart.ChartTitle.Text
        If InStr(t, "излишек") > 0 Then
            co.Chart.SeriesCollection(1).DataLabels.NumberFormat = "0"
            co.Chart.Axes(xlValue).TickLabels.NumberFormat = "0"
        ElseIf InStr(t, "перекосу") > 0 Then
            co.Chart.SeriesCollection(1).DataLabels.NumberFormat = "+0" & sep & "0%"
            co.Chart.Axes(xlValue).TickLabels.NumberFormat = "0" & sep & "0%"
        End If
    Next co
    ' заголовки рейтинга — перенос по словам
    Set pt = ws.PivotTables("П_Рейтинг")
    With pt.TableRange1.Rows(1)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 42
    End With
    On Error GoTo 0
    MsgBox "Исправления внесены.", vbInformation
End Sub

' =====================================================================
'  Дополнения: GMROI, стоимость излишка, неликвид (запускать после создания 5 новых мер)
' =====================================================================
Public Sub AddExtras()
    Dim ws As Worksheet, ptK As PivotTable, ptR As PivotTable, m, TG As String, gp As String
    Set ws = ThisWorkbook.Worksheets("Дашборд_представительства")
    Set ptK = ThisWorkbook.Worksheets("ДП_Своды").PivotTables("П_KPI")
    Set ptR = ws.PivotTables("П_Рейтинг")
    TG = ChrW(8376)
    On Error Resume Next
    For Each m In Array("GMROI", "Стоимость излишка млн", "Неликвид млн")
        ptK.CubeFields("[Measures].[" & m & "]").Orientation = xlDataField
    Next m
    gp = "'ДП_Своды'!$A$3"
    ws.Range("H13:I13").Merge
    ws.Range("J13:K13").Merge
    ws.Range("H13,J13").HorizontalAlignment = xlCenter
    ws.Range("H13").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[GMROI]""," & gp & "),"""")"
    ws.Range("H13").NumberFormat = """GMROI ""0.00"" " & TG & " маржи на 1 " & TG & " запасов"""
    ws.Range("J13").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[Неликвид млн]""," & gp & "),"""")"
    ws.Range("J13").NumberFormat = "[Red]""в т.ч. неликвид ""#,##0.0"" млн"";[Red]""в т.ч. неликвид ""#,##0.0"" млн"";""неликвида нет"""
    ws.Range("N10").Value = "Излишек сверх цели DIO, млн " & TG
    ws.Range("N13").Formula = "=IFERROR(GETPIVOTDATA(""[Measures].[Стоимость излишка млн]""," & gp & "),"""")"
    ws.Range("N13").NumberFormat = """стоит ~""#,##0.0"" млн " & TG & " в месяц"""
    ws.Range("H13,J13,N13").Font.Size = 9
    ' GMROI в рейтинг — сразу после «Маржа»
    ptR.CubeFields("[Measures].[GMROI]").Orientation = xlDataField
    With ptR.DataFields(ptR.DataFields.Count)
        .NumberFormat = "0.00"
        .Position = 6
    End With
    If Err.Number <> 0 Then
        MsgBox "Что-то не получилось: " & Err.Description, vbExclamation
    Else
        MsgBox "Готово: GMROI, стоимость излишка и неликвид добавлены.", vbInformation
    End If
End Sub
