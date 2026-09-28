Option Explicit

' ======================================================================
'  ВЫБОР ПЕРИОДА НА ОБЩЕМ ДАШБОРДЕ
'  Что делает:
'    1) добавляет в строку 5 листа «Дашборд» два поля: «Период с ... по ...»
'       со списками только закрытых месяцев;
'    2) переводит карточки на агрегат за выбранный период
'       (если выбран один месяц — значения совпадают с прежними);
'    3) переводит все графики общего дашборда на динамические диапазоны,
'       поэтому они тоже сужаются до выбранного периода;
'    4) строка изменений под карточками начинает сравнивать
'       с предыдущим периодом такой же длины.
'  Служебные ячейки: Расчет!AO1:AQ7 (правее данных, никому не мешают).
'  ПЕРЕД ЗАПУСКОМ СОХРАНИТЕ КОПИЮ ФАЙЛА.
' ======================================================================

Sub AddPeriodSelector()
    Dim wb As Workbook, wsD As Worksheet, wsR As Worksheet
    Dim nmPrefix As String

    Set wb = ActiveWorkbook
    On Error Resume Next
    Set wsD = wb.Worksheets("Дашборд")
    Set wsR = wb.Worksheets("Расчет")
    On Error GoTo 0
    If wsD Is Nothing Or wsR Is Nothing Then
        MsgBox "Не найден лист «Дашборд» или «Расчет».", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    nmPrefix = "'" & wb.Name & "'!"

    ' ---------- 1. Служебные ячейки на листе «Расчет» ----------
    With wsR
        .Range("AO1").Value = "Служебное для выбора периода на дашборде:"
        .Range("AO2").Value = "последний закрытый / конец / начало / месяцев / начало пред. периода"

        .Range("AP1").Formula = "=COUNT($W$7:$W$126)"
        .Range("AP2").Formula = "=IFERROR(MATCH(EOMONTH(Дашборд!$G$5,0),$B$7:$B$126,0),$AP$1)"
        .Range("AP3").Formula = "=MIN(IFERROR(MATCH(EOMONTH(Дашборд!$D$5,0),$B$7:$B$126,0),1),$AP$2)"
        .Range("AP4").Formula = "=$AP$2-$AP$3+1"
        .Range("AP5").Formula = "=$AP$3-$AP$4"

        ' значения предыдущего периода такой же длины
        .Range("AQ1").Formula = "=IF($AP$5>=1,IFERROR(SUM(OFFSET($J$7,$AP$5-1,0,$AP$4,1))/$AP$4*SUM(OFFSET($C$7,$AP$5-1,0,$AP$4,1))/SUM(OFFSET($G$7,$AP$5-1,0,$AP$4,1)),""""),"""")"
        .Range("AQ2").Formula = "=IF($AP$5>=1,IFERROR(SUM(OFFSET($M$7,$AP$5-1,0,$AP$4,1))/$AP$4*SUM(OFFSET($C$7,$AP$5-1,0,$AP$4,1))/SUM(OFFSET($F$7,$AP$5-1,0,$AP$4,1)),""""),"""")"
        .Range("AQ3").Formula = "=IF($AP$5>=1,IFERROR(SUM(OFFSET($P$7,$AP$5-1,0,$AP$4,1))/$AP$4*SUM(OFFSET($C$7,$AP$5-1,0,$AP$4,1))/SUM(OFFSET($R$7,$AP$5-1,0,$AP$4,1)),""""),"""")"
        .Range("AQ4").Formula = "=IF(AND(ISNUMBER($AQ$1),ISNUMBER($AQ$2)),$AQ$1+$AQ$2,"""")"
        .Range("AQ5").Formula = "=IF(AND(ISNUMBER($AQ$4),ISNUMBER($AQ$3)),$AQ$4-$AQ$3,"""")"
        .Range("AQ6").Formula = "=IF($AP$3-1>=1,IFERROR(INDEX($AF$7:$AF$126,$AP$3-1),""""),"""")"
        .Range("AQ7").Formula = "=IF($AP$3-1>=1,IFERROR(INDEX($AK$7:$AK$126,$AP$3-1),""""),"""")"
    End With

    ' ---------- 2. Именованные диапазоны ----------
    AddName wb, "rngMonthsList", "=OFFSET(Расчет!$B$7,0,0,Расчет!$AP$1,1)"
    AddName wb, "rngPeriod", RangeName("$B$7")
    AddName wb, "rngDIO", RangeName("$S$7")
    AddName wb, "rngDSO", RangeName("$T$7")
    AddName wb, "rngDPO", RangeName("$U$7")
    AddName wb, "rngCCC", RangeName("$W$7")
    AddName wb, "rngCCC_LTM", RangeName("$AF$7")
    AddName wb, "rngCHOK", RangeName("$AK$7")
    AddName wb, "rngZapasy", RangeName("$AL$7")
    AddName wb, "rngDZ", RangeName("$AM$7")
    AddName wb, "rngKZ", RangeName("$AN$7")
    AddName wb, "rngVyruchka", RangeName("$AI$7")
    AddName wb, "rngSebest", RangeName("$AJ$7")

    ' ---------- 3. Поля выбора периода в строке 5 ----------
    With wsD
        On Error Resume Next
        .Range("B5:C5").Merge
        .Range("D5:E5").Merge
        .Range("G5:H5").Merge
        .Range("J5:O5").Merge
        On Error GoTo 0

        .Range("B5").Value = "Период с:"
        .Range("B5").Font.Bold = True
        .Range("F5").Value = "по"
        .Range("F5").HorizontalAlignment = xlCenter

        .Range("J5").Value = "Выберите месяцы из списка. Для одного месяца укажите его в обоих полях. Карточки и графики пересчитаются."
        .Range("J5").Font.Italic = True
        .Range("J5").Font.Size = 9

        SetInputCell .Range("D5")
        SetInputCell .Range("G5")

        ' значения по умолчанию: весь доступный период
        .Range("D5").Value = wsR.Range("B7").Value
        .Range("G5").Value = wsR.Cells(6 + wsR.Range("AP1").Value, 2).Value
    End With

    ' ---------- 4. Карточки ----------
    With wsD
        .Range("P4").Formula = "=Расчет!$AP$2"

        .Range("B8").Formula = CardFormula("$J$7", "$G$7")
        .Range("D8").Formula = CardFormula("$M$7", "$F$7")
        .Range("F8").Formula = CardFormula("$P$7", "$R$7")
        .Range("H8").Formula = "=IFERROR($B$8+$D$8,"""")"
        .Range("J8").Formula = "=IFERROR($H$8-$F$8,"""")"
        .Range("L8").Formula = "=IFERROR(INDEX(Расчет!$AF$7:$AF$126,Расчет!$AP$2),"""")"
        .Range("N8").Formula = "=IFERROR(INDEX(Расчет!$AK$7:$AK$126,Расчет!$AP$2),"""")"

        .Range("B25").Formula = "=IF(ISNUMBER(Расчет!$AQ$1),$B$8-Расчет!$AQ$1,"""")"
        .Range("D25").Formula = "=IF(ISNUMBER(Расчет!$AQ$2),$D$8-Расчет!$AQ$2,"""")"
        .Range("F25").Formula = "=IF(ISNUMBER(Расчет!$AQ$3),$F$8-Расчет!$AQ$3,"""")"
        .Range("H25").Formula = "=IF(ISNUMBER(Расчет!$AQ$4),$H$8-Расчет!$AQ$4,"""")"
        .Range("J25").Formula = "=IF(ISNUMBER(Расчет!$AQ$5),$J$8-Расчет!$AQ$5,"""")"
        .Range("L25").Formula = "=IF(ISNUMBER(Расчет!$AQ$6),$L$8-Расчет!$AQ$6,"""")"
        .Range("N25").Formula = "=IF(ISNUMBER(Расчет!$AQ$7),$N$8-Расчет!$AQ$7,"""")"

        .Range("D11").Formula = "=IFERROR($J$8-Параметры!$C$17,"""")"
        .Range("I11").Formula = "=IFERROR(AVERAGE(OFFSET(Расчет!$W$7,Расчет!$AP$3-1,0,Расчет!$AP$4,1)),"""")"

        .Range("D4").Formula = "=INDEX(Расчет!$B$7:$B$126,Расчет!$AP$1)"
        .Range("G4").Formula = "=IF(Расчет!$AP$4=1,""Показан один месяц."",""Показан период из ""&Расчет!$AP$4&"" мес. Карточки — агрегат за период, графики сужены до него."")"

        .Range("B6").Formula = "=""КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ — ""&IF(Расчет!$AP$4=1," & _
            "TEXT(INDEX(Расчет!$B$7:$B$126,Расчет!$AP$2),""MM.YYYY"")," & _
            "TEXT(INDEX(Расчет!$B$7:$B$126,Расчет!$AP$3),""MM.YYYY"")&"" – ""&" & _
            "TEXT(INDEX(Расчет!$B$7:$B$126,Расчет!$AP$2),""MM.YYYY"")&""  ·  ""&Расчет!$AP$4&"" мес."")"

        .Range("B20").Formula = "=""СОСТАВ ЗАПАСОВ — ""&TEXT(INDEX(Расчет!$B$7:$B$126,Расчет!$AP$2),""MM.YYYY"")"
        .Range("B24").Value = "DIO на конец периода, дней"
        .Range("G11").Value = "Средний фин. цикл за период, дней"
    End With

    ' ---------- 5. Графики ----------
    RepointCharts wsD, nmPrefix

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.CalculateFull

    MsgBox "Готово." & vbCrLf & vbCrLf & _
           "В строке 5 листа «Дашборд» появились поля «Период с ... по ...»." & vbCrLf & _
           "Сейчас выбран весь доступный период." & vbCrLf & _
           "Чтобы увидеть один месяц, укажите его в обоих полях.", vbInformation
End Sub


' ---------- вспомогательные ----------

Private Function RangeName(ByVal firstCell As String) As String
    RangeName = "=OFFSET(Расчет!" & firstCell & ",Расчет!$AP$3-1,0,Расчет!$AP$4,1)"
End Function

Private Function CardFormula(ByVal baseCol As String, ByVal divCol As String) As String
    CardFormula = "=IFERROR(SUM(OFFSET(Расчет!" & baseCol & ",Расчет!$AP$3-1,0,Расчет!$AP$4,1))/Расчет!$AP$4" & _
                  "*SUM(OFFSET(Расчет!$C$7,Расчет!$AP$3-1,0,Расчет!$AP$4,1))" & _
                  "/SUM(OFFSET(Расчет!" & divCol & ",Расчет!$AP$3-1,0,Расчет!$AP$4,1)),"""")"
End Function

Private Sub AddName(wb As Workbook, ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    wb.Names(nm).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nm, RefersTo:=refersTo
End Sub

Private Sub SetInputCell(rg As Range)
    With rg
        .NumberFormat = "mmmm yyyy"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 143, 0)
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=rngMonthsList"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Месяц"
            .InputMessage = "Выберите месяц из списка закрытых месяцев."
        End With
    End With
End Sub

Private Sub RepointCharts(ws As Worksheet, ByVal nmPrefix As String)
    Dim co As ChartObject, sr As Series
    Dim f As String, nm As String, n As Long

    For Each co In ws.ChartObjects
        For Each sr In co.Chart.SeriesCollection
            f = sr.Formula
            nm = ColumnToName(f)
            If Len(nm) > 0 Then
                On Error Resume Next
                sr.XValues = "=" & nmPrefix & "rngPeriod"
                sr.Values = "=" & nmPrefix & nm
                On Error GoTo 0
                n = n + 1
            End If
        Next sr
    Next co
End Sub

Private Function ColumnToName(ByVal f As String) As String
    ' двухбуквенные колонки проверяем первыми
    If InStr(f, "Расчет!$AF$") > 0 Then ColumnToName = "rngCCC_LTM": Exit Function
    If InStr(f, "Расчет!$AK$") > 0 Then ColumnToName = "rngCHOK": Exit Function
    If InStr(f, "Расчет!$AL$") > 0 Then ColumnToName = "rngZapasy": Exit Function
    If InStr(f, "Расчет!$AM$") > 0 Then ColumnToName = "rngDZ": Exit Function
    If InStr(f, "Расчет!$AN$") > 0 Then ColumnToName = "rngKZ": Exit Function
    If InStr(f, "Расчет!$AI$") > 0 Then ColumnToName = "rngVyruchka": Exit Function
    If InStr(f, "Расчет!$AJ$") > 0 Then ColumnToName = "rngSebest": Exit Function
    If InStr(f, "Расчет!$S$") > 0 Then ColumnToName = "rngDIO": Exit Function
    If InStr(f, "Расчет!$T$") > 0 Then ColumnToName = "rngDSO": Exit Function
    If InStr(f, "Расчет!$U$") > 0 Then ColumnToName = "rngDPO": Exit Function
    If InStr(f, "Расчет!$W$") > 0 Then ColumnToName = "rngCCC": Exit Function
    ColumnToName = ""
End Function


' ======================================================================
'  ОТКАТ: вернуть дашборд к прежнему виду (карточки = последний месяц,
'  графики = вся история). Служебные ячейки при этом очищаются.
' ======================================================================
Sub RemovePeriodSelector()
    Dim wb As Workbook, wsD As Worksheet, wsR As Worksheet
    Dim co As ChartObject, sr As Series
    Dim nm As String, col As String

    Set wb = ActiveWorkbook
    Set wsD = wb.Worksheets("Дашборд")
    Set wsR = wb.Worksheets("Расчет")

    Application.ScreenUpdating = False

    With wsD
        .Range("P4").Formula = "=COUNT(Расчет!$B$7:$B$126)"
        .Range("B8").Formula = "=IFERROR(INDEX(Расчет!$S$7:$S$126,$P$4),"""")"
        .Range("D8").Formula = "=IFERROR(INDEX(Расчет!$T$7:$T$126,$P$4),"""")"
        .Range("F8").Formula = "=IFERROR(INDEX(Расчет!$U$7:$U$126,$P$4),"""")"
        .Range("H8").Formula = "=IFERROR(INDEX(Расчет!$V$7:$V$126,$P$4),"""")"
        .Range("J8").Formula = "=IFERROR(INDEX(Расчет!$W$7:$W$126,$P$4),"""")"
        .Range("L8").Formula = "=IFERROR(INDEX(Расчет!$AF$7:$AF$126,$P$4),"""")"
        .Range("N8").Formula = "=IFERROR(INDEX(Расчет!$AK$7:$AK$126,$P$4),"""")"
        .Range("D11").Formula = "=IFERROR(INDEX(Расчет!$W$7:$W$126,$P$4)-Параметры!$C$17,"""")"
        .Range("D4").Formula = "=INDEX(Расчет!$B$7:$B$126,$P$4)"
        .Range("B6").Value = "КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ — ПОСЛЕДНИЙ ЗАКРЫТЫЙ МЕСЯЦ"
        .Range("B20").Value = "СОСТАВ ЗАПАСОВ — ПОСЛЕДНИЙ ЗАКРЫТЫЙ МЕСЯЦ"
        .Range("B24").Value = "DIO за месяц, дней"

        RestoreDelta .Range("B25"), "$S$"
        RestoreDelta .Range("D25"), "$T$"
        RestoreDelta .Range("F25"), "$U$"
        RestoreDelta .Range("H25"), "$V$"
        RestoreDelta .Range("J25"), "$W$"
        RestoreDelta .Range("L25"), "$AF$"
        RestoreDelta .Range("N25"), "$AK$"

        On Error Resume Next
        .Range("D5").Validation.Delete
        .Range("G5").Validation.Delete
        .Range("B5:O5").ClearContents
        .Range("B5:O5").Interior.Pattern = xlNone
        On Error GoTo 0
    End With

    For Each co In wsD.ChartObjects
        For Each sr In co.Chart.SeriesCollection
            nm = ColumnToName(sr.Formula)
            col = NameToColumn(sr.Formula)
            If Len(col) > 0 Then
                On Error Resume Next
                sr.XValues = wsR.Range("$B$7:$B$126")
                sr.Values = wsR.Range(col & "7:" & col & "126")
                On Error GoTo 0
            End If
        Next sr
    Next co

    wsR.Range("AO1:AQ7").ClearContents

    Application.ScreenUpdating = True
    Application.CalculateFull
    MsgBox "Дашборд возвращён к прежнему виду.", vbInformation
End Sub

Private Sub RestoreDelta(rg As Range, ByVal col As String)
    rg.Formula = "=IF($P$4>1,IFERROR(INDEX(Расчет!" & col & "7:" & col & "126,$P$4)" & _
                 "-INDEX(Расчет!" & col & "7:" & col & "126,$P$4-1),""""),"""")"
End Sub

Private Function NameToColumn(ByVal f As String) As String
    If InStr(f, "rngCCC_LTM") > 0 Then NameToColumn = "AF": Exit Function
    If InStr(f, "rngCHOK") > 0 Then NameToColumn = "AK": Exit Function
    If InStr(f, "rngZapasy") > 0 Then NameToColumn = "AL": Exit Function
    If InStr(f, "rngDZ") > 0 Then NameToColumn = "AM": Exit Function
    If InStr(f, "rngKZ") > 0 Then NameToColumn = "AN": Exit Function
    If InStr(f, "rngVyruchka") > 0 Then NameToColumn = "AI": Exit Function
    If InStr(f, "rngSebest") > 0 Then NameToColumn = "AJ": Exit Function
    If InStr(f, "rngDIO") > 0 Then NameToColumn = "S": Exit Function
    If InStr(f, "rngDSO") > 0 Then NameToColumn = "T": Exit Function
    If InStr(f, "rngDPO") > 0 Then NameToColumn = "U": Exit Function
    If InStr(f, "rngCCC") > 0 Then NameToColumn = "W": Exit Function
    NameToColumn = ""
End Function
