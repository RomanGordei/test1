Attribute VB_Name = "X5Report"
'==============================================================================
'  X5 Dialog · сборщик отчёта на модели куба
'  Прототип 0.1 · месячный куб · листы «Обложка», «Резюме», «Категории», «Проверка»
'
'  Как запустить
'    1. Откройте файл куба (X5_Cube_monthly.xlsx) и книгу со сборщиком.
'    2. Перейдите в файл куба и запустите макрос СобратьОтчёт (Вид > Макросы).
'
'  Что делает сборщик
'    • Добавляет в модель таблицы выбора «Окно», «Сравнение», «Единица» и
'      таблицу «Отчёт», к которой привязаны его меры (все называются «x5 ...»).
'    • Строит листы отчёта: срезы, плитки, графики, сводную и проверку.
'    • Ваши листы и меры не трогает. Повторный запуск пересобирает только своё.
'==============================================================================
Option Explicit

' ---- Настройки ---------------------------------------------------------------
Private Const CN As String = "ThisWorkbookDataModel"
Private Const OWN As String = "Маревен Фуд Сэнтрал"             ' свой производитель
Private Const COL_PROD As String = "Производитель (корректный)"  ' столбец производителя в «Продукты»
Private Const T_REPORT As String = "Отчёт"                       ' таблица модели для мер сборщика
Private Const SH_SET As String = "Настройки отчёта"
Private Const SH_LOG As String = "Журнал сборки"
Private Const SH_COVER As String = "Обложка"
Private Const SH_SUM As String = "Резюме"
Private Const SH_CAT As String = "Категории"
Private Const SH_CHK As String = "Проверка"
Private Const MARK As String = "x5report"                        ' метка листов сборщика
Private Const CACHES As String = "srPeriod srWindow srCompare srUnit srNet srFO srRegion srCat"

' ---- Константы Excel: свои, чтобы не зависеть от библиотеки типов ------------
Private Const XL_EXTERNAL As Long = 2          ' xlExternal
Private Const XL_PT_V15 As Long = 5            ' xlPivotTableVersion15
Private Const XL_ROW_FIELD As Long = 1         ' xlRowField
Private Const XL_TABULAR As Long = 1           ' xlTabularRow
Private Const XL_DESCENDING As Long = 2        ' xlDescending
Private Const XL_SORT_SOURCE As Long = 1       ' xlSlicerSortDataSourceOrder
Private Const XL_SORT_DESC As Long = 3         ' xlSlicerSortDescending
Private Const XL_SRC_RANGE As Long = 1         ' xlSrcRange
Private Const XL_YES As Long = 1               ' xlYes
Private Const XL_CMD_EXCEL As Long = 7         ' xlCmdExcel
Private Const XL_LINE As Long = 4              ' xlLine
Private Const XL_BAR As Long = 57              ' xlBarClustered
Private Const XL_CATEGORY As Long = 1          ' xlCategory
Private Const XL_VALUE As Long = 2             ' xlValue
Private Const XL_LEFT As Long = -4131          ' xlLeft
Private Const XL_VCENTER As Long = -4108       ' xlCenter
Private Const XL_TOP As Long = -4160           ' xlTop и xlLegendPositionTop
Private Const XL_TICK_LOW As Long = -4134      ' xlTickLabelPositionLow
Private Const XL_LABEL_OUT As Long = 2         ' xlLabelPositionOutsideEnd
Private Const XL_SHEET_HIDDEN As Long = 0      ' xlSheetHidden
Private Const XL_SHEET_VISIBLE As Long = -1    ' xlSheetVisible
Private Const XL_EDGE_LEFT As Long = 7         ' xlEdgeLeft
Private Const XL_CONTINUOUS As Long = 1        ' xlContinuous
Private Const XL_THICK As Long = 4             ' xlThick
Private Const XL_CELL_VALUE As Long = 1        ' xlCellValue
Private Const XL_EQUAL As Long = 3             ' xlEqual
Private Const XL_NOT_EQUAL As Long = 4         ' xlNotEqual
Private Const MSO_FALSE As Long = 0            ' msoFalse
Private Const MSO_DASH As Long = 4             ' msoLineDash

' ---- Состояние сборки ----------------------------------------------------------
Private wb As Object                ' книга куба, в которой строится отчёт
Private wsLog As Object
Private logRow As Long
Private known As Object             ' имена мер, которые уже есть в модели
Private srOk As Object              ' срезы, которые удалось создать (имя кэша -> True)
Private scWay As Long               ' способ создания среза, который сработал
Private fDec0 As Object, fDec1 As Object, fPct1 As Object, fInt As Object, fCode As Object, fGen As Object
Private Q As String, UPA As String, DNA As String, ARW As String
Private ALL_SR As String, MKT_SR As String
Private F_NUM1 As String, F_UP_PCT As String, F_UP_PP As String, F_PCT As String, F_PP As String
Private F_PTS As String, F_INT As String, F_PIV_UP_PCT As String, F_PIV_UP As String, F_PIV_PCT As String, F_PIV_PP As String
Private C_ACCENT As Long, C_OTHER As Long, C_GOOD As Long, C_BAD As Long
Private C_INK As Long, C_INK2 As Long, C_MUTED As Long, C_GRID As Long, C_TILE As Long

'==============================================================================
'  Макросы для запуска
'==============================================================================

' Собирает (или пересобирает) отчёт в активной книге куба
Public Sub СобратьОтчёт()
    Dim stepName As String, failMsg As String, ok As Boolean, msg As String, miss As String
    If Not PickWorkbook() Then Exit Sub
    InitConsts
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    stepName = "журнал": StartLog
    stepName = "проверка модели": CheckModel
    stepName = "удаление прошлой сборки": RemoveOldBuild
    stepName = "таблицы выбора": BuildSettings
    stepName = "меры": BuildMeasures
    stepName = "лист «" & SH_SUM & "»": BuildSummary
    stepName = "лист «" & SH_CAT & "»": BuildCategories
    stepName = "лист «" & SH_COVER & "»": BuildCover
    stepName = "лист «" & SH_CHK & "»": BuildChecks
    stepName = "порядок листов": OrderSheets
    stepName = "расчёт формул"
    WriteLog "Считаю формулы отчёта"
    Application.Calculate
    Application.CalculateUntilAsyncQueriesDone
    ok = True
Finish:
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If ok Then
        WriteLog "Готово. Проверки: " & CountOk() & " из 6 OK."
        wb.Worksheets(SH_COVER).Activate
        msg = "Отчёт собран. Проверки: " & CountOk() & " из 6 OK (лист «" & SH_CHK & "»)."
        miss = MissingSlicers()
        If Len(miss) > 0 Then
            msg = msg & vbLf & vbLf & "Не создались срезы: " & miss & ". Вместо них отчёт берёт значения по умолчанию " & _
                  "(последний месяц, MAT к YA, рубли, весь рынок). Подробности на листе «" & SH_LOG & "»."
            MsgBox msg, vbExclamation, "Сборка отчёта"
        Else
            MsgBox msg, vbInformation, "Сборка отчёта"
        End If
    Else
        WriteLog failMsg
        If Not wsLog Is Nothing Then wsLog.Activate
        MsgBox failMsg & vbLf & vbLf & "Подробности на листе «" & SH_LOG & "».", vbCritical, "Сборка отчёта"
    End If
    Exit Sub
Fail:
    failMsg = "Сборка остановилась на шаге «" & stepName & "»: " & Err.Description
    Resume Finish
End Sub

' Обновляет модель и формулы отчёта, показывает итог проверки
Public Sub ОбновитьОтчёт()
    Dim n As Long
    If Not PickWorkbook() Then Exit Sub
    InitConsts
    Application.StatusBar = "Обновляю модель и отчёт" & ChrW(8230)
    wb.RefreshAll
    Application.Calculate
    Application.CalculateUntilAsyncQueriesDone
    Application.StatusBar = False
    On Error Resume Next
    wb.Worksheets(SH_COVER).Range("B4").Value = "Обновлено: " & Format(Now, "dd.mm.yyyy hh:nn")
    On Error GoTo 0
    n = CountOk()
    MsgBox "Модель и отчёт обновлены. Проверки: " & n & " из 6 OK.", IIf(n = 6, vbInformation, vbExclamation), "Отчёт X5"
End Sub

' Убирает из книги всё, что добавил сборщик: листы, срезы, меры и таблицы выбора
Public Sub УдалитьОтчёт()
    Dim nm As Variant, ws As Object
    If Not PickWorkbook() Then Exit Sub
    If MsgBox("Удалить из книги «" & wb.Name & "» листы, срезы, меры и таблицы выбора отчёта?", _
              vbQuestion + vbYesNo, "Отчёт X5") <> vbYes Then Exit Sub
    InitConsts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    RemoveOldBuild
    On Error Resume Next
    For Each nm In Array("Окно", "Сравнение", "Единица", T_REPORT)
        wb.Connections("WorksheetConnection_" & nm).Delete
    Next
    For Each nm In Array(SH_SET, SH_LOG)
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(nm))
        If Not ws Is Nothing Then
            ws.Visible = XL_SHEET_VISIBLE
            ws.Delete
        End If
    Next
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Отчёт удалён из книги.", vbInformation, "Отчёт X5"
End Sub

'==============================================================================
'  Подготовка
'==============================================================================

Private Function PickWorkbook() As Boolean
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    If Not HasModelTable("Факт") Or Not HasModelTable("Календарь") Then
        MsgBox "Сделайте активным файл куба: в книге «" & wb.Name & "» нет модели с таблицами «Факт» и «Календарь».", _
               vbExclamation, "Отчёт X5"
        Exit Function
    End If
    If HasModelColumn("Календарь", "НеделяМетка") Or HasModelColumn("Календарь", "Неделя") Or _
       InStr(1, wb.Name, "weekly", vbTextCompare) > 0 Then
        MsgBox "Это недельный куб. Прототип пока собирается на месячном.", vbInformation, "Отчёт X5"
        Exit Function
    End If
    PickWorkbook = True
End Function

Private Sub InitConsts()
    Q = Chr(34): UPA = ChrW(9650): DNA = ChrW(9660): ARW = ChrW(8594)
    Set srOk = CreateObject("Scripting.Dictionary")
    scWay = 0
    ALL_SR = "srPeriod,srWindow,srCompare,srUnit,srNet,srFO,srRegion,srCat"
    MKT_SR = "srNet,srFO,srRegion,srCat"
    ' форматы плиток: стрелка и цвет для роста и доли, нейтральный знак для цен и промо
    F_NUM1 = "#,##0.0"
    F_UP_PCT = "[Color10]" & Q & UPA & " " & Q & "0.0%;[Red]" & Q & DNA & " " & Q & "0.0%;0.0%"
    F_UP_PP = "[Color10]" & Q & UPA & " " & Q & "0.0" & Q & " п.п." & Q & ";[Red]" & Q & DNA & " " & Q & _
              "0.0" & Q & " п.п." & Q & ";0.0" & Q & " п.п." & Q
    F_PCT = "+0.0%;-0.0%;0.0%"
    F_PP = "+0.0" & Q & " п.п." & Q & ";-0.0" & Q & " п.п." & Q & ";0.0" & Q & " п.п." & Q
    F_PTS = "+0;-0;0"
    F_INT = "+#,##0;-#,##0;0"
    ' форматы столбцов сводной: со знаком, без стрелок
    F_PIV_UP_PCT = "[Color10]+0.0%;[Red]-0.0%;0.0%"
    F_PIV_UP = "[Color10]+0.0;[Red]-0.0;0.0"
    F_PIV_PCT = "+0.0%;-0.0%;0.0%"
    F_PIV_PP = "+0.0;-0.0;0.0"
    ' палитра та же, что в HTML-дашборде
    C_ACCENT = RGB(42, 120, 214): C_OTHER = RGB(196, 194, 188)
    C_GOOD = RGB(0, 99, 0): C_BAD = RGB(208, 59, 59)
    C_INK = RGB(11, 11, 11): C_INK2 = RGB(82, 81, 78): C_MUTED = RGB(137, 135, 129)
    C_GRID = RGB(232, 231, 225): C_TILE = RGB(245, 245, 242)
End Sub

Private Function HasModelTable(ByVal nm As String) As Boolean
    Dim t As Object
    On Error GoTo NoTable
    For Each t In wb.Model.ModelTables
        If t.Name = nm Then HasModelTable = True: Exit Function
    Next
NoTable:
End Function

Private Function HasModelColumn(ByVal tbl As String, ByVal col As String) As Boolean
    Dim t As Object, c As Object
    On Error GoTo NoColumn
    For Each t In wb.Model.ModelTables
        If t.Name = tbl Then
            For Each c In t.ModelTableColumns
                If c.Name = col Then HasModelColumn = True: Exit Function
            Next
        End If
    Next
NoColumn:
End Function

Private Sub CheckModel()
    ' Проверяются только поля, которые нужны срезам и сводным: они должны быть видимыми.
    ' Скрытые столбцы (Факт, итоги категории, НачалоПериода) Excel здесь не показывает, но меры их видят.
    ' Если какого-то из них нет, сборка остановится на создании меры и покажет её формулу.
    Dim need As Variant, parts() As String, missing As String, i As Long
    need = Array("Календарь|Период", "Календарь|МесяцМетка", "Продукты|Категория", "Продукты|" & COL_PROD, _
                 "Регионы|ФО", "Регионы|СубъектРФ", "Сети|Сеть")
    For i = LBound(need) To UBound(need)
        parts = Split(need(i), "|")
        If Not HasModelColumn(parts(0), parts(1)) Then missing = missing & ", " & parts(0) & "[" & parts(1) & "]"
    Next
    If Len(missing) > 0 Then Err.Raise vbObjectError + 600, , "в модели нет видимых столбцов: " & Mid$(missing, 3) & _
        ". Если столбец скрыт в Power Pivot, покажите его: правый клик по столбцу > Показать в клиентских средствах"
    WriteLog "Модель на месте: поля для срезов и сводных найдены"
End Sub

Private Sub StartLog()
    Set wsLog = Nothing
    On Error Resume Next
    Set wsLog = wb.Worksheets(SH_LOG)
    On Error GoTo 0
    If wsLog Is Nothing Then
        Set wsLog = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsLog.Name = SH_LOG
        wsLog.CustomProperties.Add MARK, "log"
    End If
    wsLog.Cells.Clear
    wsLog.Columns(1).ColumnWidth = 10
    wsLog.Columns(2).ColumnWidth = 150
    wsLog.Range("A1").Value = "Журнал сборки отчёта"
    wsLog.Range("A1").Font.Bold = True
    logRow = 2
    WriteLog "Книга: " & wb.Name
End Sub

Private Sub WriteLog(ByVal text As String)
    If wsLog Is Nothing Then Exit Sub
    logRow = logRow + 1
    wsLog.Cells(logRow, 1).Value = Format(Now, "hh:nn:ss")
    wsLog.Cells(logRow, 2).Value = "'" & text
    On Error Resume Next
    Application.StatusBar = Left$(Replace(Replace(text, vbCr, " "), vbLf, " "), 200)
    On Error GoTo 0
End Sub

Private Sub Status(ByVal text As String)
    On Error Resume Next
    Application.StatusBar = Left$(text, 200)
    On Error GoTo 0
End Sub

Private Function CountOk() As Long
    On Error Resume Next
    CountOk = Application.WorksheetFunction.CountIf(wb.Worksheets(SH_CHK).Range("E5:E10"), "OK")
End Function

Private Function IsOurSheet(ByVal ws As Object) As Boolean
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).Name = MARK Then IsOurSheet = True: Exit Function
    Next
End Function

' Удаляет листы, срезы и меры прошлой сборки. Лист с таблицами выбора и журнал остаются.
Private Sub RemoveOldBuild()
    Dim i As Long, n As Long, ws As Object, xb As Excel.Workbook
    Set xb = wb
    For i = xb.SlicerCaches.Count To 1 Step -1
        If InStr(" " & CACHES & " ", " " & xb.SlicerCaches(i).Name & " ") > 0 Then xb.SlicerCaches(i).Delete
    Next
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If IsOurSheet(ws) And ws.Name <> SH_SET And ws.Name <> SH_LOG Then ws.Delete
    Next
    If HasModelTable(T_REPORT) Then
        For i = wb.Model.ModelMeasures.Count To 1 Step -1
            If wb.Model.ModelMeasures(i).AssociatedTable.Name = T_REPORT Then
                wb.Model.ModelMeasures(i).Delete
                n = n + 1
            End If
        Next
    End If
    WriteLog "Прошлая сборка убрана, удалено мер: " & n
End Sub

'==============================================================================
'  Таблицы выбора в модели
'==============================================================================

Private Sub BuildSettings()
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(SH_SET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SH_SET
        ws.CustomProperties.Add MARK, "settings"
        ws.Range("A1").Value = "Таблицы выбора для срезов отчёта. Не переименовывайте и не удаляйте их: на них ссылаются меры."
    End If
    EnsureList ws, "Окно", "A3", Rows2(Array("Окно", "Порядок"), Array("Месяц", 1), Array("L3M", 2), _
                                     Array("YTD", 3), Array("MAT", 4))
    EnsureList ws, "Сравнение", "D3", Rows2(Array("Сравнение", "Порядок"), Array("YA", 1), Array("PP", 2))
    EnsureList ws, "Единица", "G3", Rows2(Array("Единица", "Порядок"), Array("руб", 1), Array("кг", 2), Array("шт", 3))
    EnsureList ws, T_REPORT, "J3", Rows2(Array("Меры отчёта"), Array("меры x5"))
    AddToModel ws, "Окно"
    AddToModel ws, "Сравнение"
    AddToModel ws, "Единица"
    AddToModel ws, T_REPORT
    ws.Visible = XL_SHEET_HIDDEN
End Sub

' Строки в двумерный массив для записи на лист
Private Function Rows2(ParamArray rowsIn() As Variant) As Variant
    Dim out() As Variant, r As Long, c As Long, nR As Long, nC As Long, first As Variant
    nR = UBound(rowsIn) - LBound(rowsIn) + 1
    first = rowsIn(LBound(rowsIn))
    nC = UBound(first) - LBound(first) + 1
    ReDim out(1 To nR, 1 To nC)
    For r = 1 To nR
        For c = 1 To nC
            out(r, c) = rowsIn(LBound(rowsIn) + r - 1)(LBound(first) + c - 1)
        Next
    Next
    Rows2 = out
End Function

Private Sub EnsureList(ByVal ws As Object, ByVal nm As String, ByVal anchor As String, ByVal data As Variant)
    Dim lo As Object, rng As Object
    On Error Resume Next
    Set lo = ws.ListObjects(nm)
    On Error GoTo 0
    If Not lo Is Nothing Then Exit Sub
    Set rng = ws.Range(anchor).Resize(UBound(data, 1), UBound(data, 2))
    rng.Value = data
    Set lo = ws.ListObjects.Add(XL_SRC_RANGE, rng, , XL_YES)
    lo.Name = nm
End Sub

Private Sub AddToModel(ByVal ws As Object, ByVal nm As String)
    Dim tries As Variant, i As Long, con As String, d As String
    If HasModelTable(nm) Then Exit Sub
    con = "WorksheetConnection_" & nm
    tries = Array(Array("WORKSHEET;" & wb.FullName, wb.Name & "!" & nm), _
                  Array("WORKSHEET;" & wb.Path & Application.PathSeparator & "[" & wb.Name & "]" & ws.Name, wb.Name & "!" & nm), _
                  Array("WORKSHEET;" & wb.FullName, nm))
    For i = 0 To UBound(tries)
        On Error Resume Next
        wb.Connections(con).Delete
        Err.Clear
        wb.Connections.Add2 con, "", tries(i)(0), tries(i)(1), XL_CMD_EXCEL, True, False
        d = Err.Description
        On Error GoTo 0
        If HasModelTable(nm) Then
            WriteLog "В модель добавлена таблица «" & nm & "»"
            Exit Sub
        End If
        WriteLog "Попытка " & (i + 1) & " добавить «" & nm & "» в модель не прошла: " & d
    Next
    On Error Resume Next
    wb.Connections(con).Delete
    On Error GoTo 0
    Err.Raise vbObjectError + 602, , "не удалось добавить таблицу «" & nm & "» в модель (" & d & "). " & _
        "Добавьте её вручную: покажите лист «" & SH_SET & "», щёлкните в таблицу «" & nm & _
        "», Power Pivot > Добавить в модель данных. Затем запустите сборку снова."
End Sub

'==============================================================================
'  Меры
'==============================================================================

Private Sub BuildMeasures()
    Dim i As Long, c As Variant, own_ As String
    Set known = CreateObject("Scripting.Dictionary")
    known.CompareMode = 1
    For i = 1 To wb.Model.ModelMeasures.Count
        known(wb.Model.ModelMeasures(i).Name) = True
    Next
    Set fDec0 = wb.Model.ModelFormatDecimalNumber(True, 0)
    Set fDec1 = wb.Model.ModelFormatDecimalNumber(True, 1)
    Set fPct1 = wb.Model.ModelFormatPercentageNumber(False, 1)
    Set fInt = wb.Model.ModelFormatWholeNumber(True)
    Set fCode = wb.Model.ModelFormatWholeNumber(False)
    Set fGen = wb.Model.ModelFormatGeneral
    own_ = "'Продукты'[" & COL_PROD & "] = """ & OWN & """"

    ' --- Выбор в срезах: последний месяц, окно, сравнение, единица ---------------
    AddM "x5_Конец", "MAX('Календарь'[НачалоПериода])", fGen
    AddM "x5_Окно", "IF(HASONEVALUE('Окно'[Окно]), VALUES('Окно'[Окно]), ""MAT"")", fGen
    AddM "x5_Сравнение", "IF(HASONEVALUE('Сравнение'[Сравнение]), VALUES('Сравнение'[Сравнение]), ""YA"")", fGen
    AddM "x5_Единица", "IF(HASONEVALUE('Единица'[Единица]), VALUES('Единица'[Единица]), ""руб"")", fGen
    AddM "x5_Начало", "VAR vEnd = [x5_Конец] RETURN SWITCH([x5_Окно], ""Месяц"", vEnd, ""L3M"", EDATE(vEnd, -2), " & _
                      """YTD"", DATE(YEAR(vEnd), 1, 1), EDATE(vEnd, -11))", fGen
    AddM "x5_Сдвиг", "VAR vEnd = [x5_Конец] VAR vStart = [x5_Начало] RETURN IF([x5_Сравнение] = ""PP"", " & _
                     "(YEAR(vEnd) - YEAR(vStart)) * 12 + MONTH(vEnd) - MONTH(vStart) + 1, 12)", fGen
    AddM "x5_Масштаб", "SWITCH([x5_Единица], ""кг"", 1000, 1000000)", fGen
    AddM "x5_Подпись единицы", "SWITCH([x5_Единица], ""кг"", ""т"", ""шт"", ""млн шт"", ""млн руб"")", fGen
    AddM "x5_Конец ГГГГММ", "VAR vEnd = [x5_Конец] RETURN YEAR(vEnd) * 100 + MONTH(vEnd)", fCode
    AddM "x5_Первый ГГГГММ", "VAR vD = CALCULATE(MIN('Календарь'[НачалоПериода]), ALL('Календарь')) " & _
                             "RETURN YEAR(vD) * 100 + MONTH(vD)", fCode
    AddM "x5_Последний ГГГГММ", "VAR vD = CALCULATE(MAX('Календарь'[НачалоПериода]), ALL('Календарь')) " & _
                                "RETURN YEAR(vD) * 100 + MONTH(vD)", fCode

    ' --- Суммы за окно (_w) и за сравнение (_c) ----------------------------------
    For Each c In Array("TOTAL_Руб", "TOTAL_Шт", "TOTAL_КгЛ", "PROMO_Руб", "PROMO_Шт", "PROMO_КгЛ", _
                        "LOYAL_Руб", "LOYAL_Шт", "LOYAL_КгЛ")
        AddM "x5_w " & c, Win("SUM('Факт'[" & c & "])"), fDec0
        AddM "x5_c " & c, Cmp("SUM('Факт'[" & c & "])"), fDec0
    Next
    AddM "x5_w объём", UnitPick("x5_w TOTAL"), fDec0
    AddM "x5_c объём", UnitPick("x5_c TOTAL"), fDec0
    AddM "x5_w промо", UnitPick("x5_w PROMO"), fDec0
    AddM "x5_c промо", UnitPick("x5_c PROMO"), fDec0
    AddM "x5_w карты", UnitPick("x5_w LOYAL"), fDec0
    AddM "x5_c карты", UnitPick("x5_c LOYAL"), fDec0
    ' знаменатели: категории из среза (без фильтров производителя, бренда, SKU) и вся ПБП
    AddM "x5_w объём кат", "CALCULATE([x5_w объём], ALLEXCEPT('Продукты', 'Продукты'[Категория]))", fDec0
    AddM "x5_c объём кат", "CALCULATE([x5_c объём], ALLEXCEPT('Продукты', 'Продукты'[Категория]))", fDec0
    AddM "x5_w объём ПБП", "CALCULATE([x5_w объём], ALL('Продукты'))", fDec0
    AddM "x5_c объём ПБП", "CALCULATE([x5_c объём], ALL('Продукты'))", fDec0
    AddM "x5_w объём Маревен", "CALCULATE([x5_w объём], " & own_ & ")", fDec0
    AddM "x5_c объём Маревен", "CALCULATE([x5_c объём], " & own_ & ")", fDec0
    ' TDP месяца = магазины SKU / магазины категории; за окно — среднее по месяцам, где категория продавалась
    AddM "x5_TDP мес", "VAR vC = SUM('Категория'[TOTAL_Магазины]) " & _
                       "RETURN IF(vC > 0, DIVIDE(SUM('Факт'[TOTAL_Магазины]), vC) + 0)", fDec1
    AddM "x5_w TDP", VarsWin() & "RETURN AVERAGEX(" & RangeWin() & ", [x5_TDP мес])", fDec1
    AddM "x5_c TDP", VarsCmp() & "RETURN AVERAGEX(" & RangeCmp() & ", [x5_TDP мес])", fDec1
    AddM "x5_Руб мес", "SUM('Факт'[TOTAL_Руб])", fDec0

    ' --- Продажи, прирост, доли ---------------------------------------------------
    AddM "x5 Продажи", "DIVIDE([x5_w объём], [x5_Масштаб])", fDec1
    AddM "x5 Продажи сравн", "DIVIDE([x5_c объём], [x5_Масштаб])", fDec1
    AddM "x5 Прирост", "VAR vC = [x5_c объём] RETURN IF(vC > 0, DIVIDE([x5_w объём], vC) - 1)", fPct1
    AddM "x5 Доля кат", "DIVIDE([x5_w объём], [x5_w объём кат])", fPct1
    AddM "x5 Доля кат сравн", "DIVIDE([x5_c объём], [x5_c объём кат])", fPct1
    AddM "x5 Изм доли кат", "VAR vB = [x5 Доля кат сравн] RETURN IF(NOT ISBLANK(vB), ([x5 Доля кат] - vB) * 100)", fDec1
    AddM "x5 Доля ПБП", "DIVIDE([x5_w объём], [x5_w объём ПБП])", fPct1
    AddM "x5 Изм доли ПБП", "VAR vB = DIVIDE([x5_c объём], [x5_c объём ПБП]) " & _
                            "RETURN IF(NOT ISBLANK(vB), ([x5 Доля ПБП] - vB) * 100)", fDec1
    AddM "x5 Продажи Маревен", "DIVIDE([x5_w объём Маревен], [x5_Масштаб])", fDec1
    AddM "x5 Прирост Маревен", "VAR vC = [x5_c объём Маревен] RETURN IF(vC > 0, DIVIDE([x5_w объём Маревен], vC) - 1)", fPct1
    AddM "x5 Доля Маревен", "DIVIDE([x5_w объём Маревен], [x5_w объём кат])", fPct1
    AddM "x5 Изм доли Маревен", "VAR vB = DIVIDE([x5_c объём Маревен], [x5_c объём кат]) " & _
                                "RETURN IF(NOT ISBLANK(vB), ([x5 Доля Маревен] - vB) * 100)", fDec1

    ' --- Цены и индекс цены (за штуку, а при единице «кг» — за килограмм) ------------
    AddM "x5 Цена кг", "DIVIDE([x5_w TOTAL_Руб], [x5_w TOTAL_КгЛ])", fDec1
    AddM "x5 Изм цены кг", "VAR vB = DIVIDE([x5_c TOTAL_Руб], [x5_c TOTAL_КгЛ]) " & _
                           "RETURN IF(vB > 0, DIVIDE([x5 Цена кг], vB) - 1)", fPct1
    AddM "x5 Цена шт", "DIVIDE([x5_w TOTAL_Руб], [x5_w TOTAL_Шт])", fDec1
    AddM "x5 Изм цены шт", "VAR vB = DIVIDE([x5_c TOTAL_Руб], [x5_c TOTAL_Шт]) " & _
                           "RETURN IF(vB > 0, DIVIDE([x5 Цена шт], vB) - 1)", fPct1
    AddM "x5 Индекс цены", PriceIndex("x5_w"), fDec0
    AddM "x5 Индекс цены сравн", PriceIndex("x5_c"), fDec0
    AddM "x5 Изм индекса", "VAR vA = [x5 Индекс цены] VAR vB = [x5 Индекс цены сравн] " & _
                           "RETURN IF(NOT ISBLANK(vA) && NOT ISBLANK(vB), vA - vB)", fDec0
    AddM "x5 Индекс цены Маревен", "CALCULATE([x5 Индекс цены], " & own_ & ")", fDec0
    AddM "x5 Изм индекса Маревен", "CALCULATE([x5 Изм индекса], " & own_ & ")", fDec0

    ' --- Промо и карты лояльности ------------------------------------------------------
    AddM "x5 Доля промо", "DIVIDE([x5_w промо], [x5_w объём])", fPct1
    AddM "x5 Изм доли промо", "VAR vB = DIVIDE([x5_c промо], [x5_c объём]) " & _
                              "RETURN IF(NOT ISBLANK(vB), ([x5 Доля промо] - vB) * 100)", fDec1
    AddM "x5 Доля промо Маревен", "CALCULATE([x5 Доля промо], " & own_ & ")", fPct1
    AddM "x5 Изм доли промо Маревен", "CALCULATE([x5 Изм доли промо], " & own_ & ")", fDec1
    AddM "x5 Доля карты", "DIVIDE([x5_w карты], [x5_w объём])", fPct1
    AddM "x5 Изм доли карты", "VAR vB = DIVIDE([x5_c карты], [x5_c объём]) " & _
                              "RETURN IF(NOT ISBLANK(vB), ([x5 Доля карты] - vB) * 100)", fDec1

    ' --- Ассортимент и полка ----------------------------------------------------------
    AddM "x5 SKU", VarsWin() & "RETURN CALCULATE(DISTINCTCOUNT('Факт'[PLU]), " & RangeWin() & ", 'Факт'[TOTAL_Шт] > 0)", fInt
    AddM "x5 SKU сравн", VarsCmp() & "RETURN CALCULATE(DISTINCTCOUNT('Факт'[PLU]), " & RangeCmp() & ", 'Факт'[TOTAL_Шт] > 0)", fInt
    AddM "x5 Изм SKU", "VAR vB = [x5 SKU сравн] RETURN IF(NOT ISBLANK(vB), [x5 SKU] - vB)", fInt
    AddM "x5 TDP", "[x5_w TDP]", fDec1
    AddM "x5 Изм TDP", "VAR vB = [x5_c TDP] RETURN IF(vB > 0, DIVIDE([x5_w TDP], vB) - 1)", fPct1

    ' --- Разложение изменения продаж в рублях: полка x спрос на точку x вес x цена за кг ---
    ' Вклады считаются через логарифмическое среднее и в сумме дают ровно изменение продаж.
    AddM "x5_L", "VAR vR1 = [x5_w TOTAL_Руб] VAR vR0 = [x5_c TOTAL_Руб] RETURN IF(vR1 > 0 && vR0 > 0, " & _
                 "IF(vR1 = vR0, vR1, (vR1 - vR0) / (LN(vR1) - LN(vR0))))", fDec0
    AddM "x5 Вклад полки", Contrib("[x5_w TDP]", "[x5_c TDP]"), fDec1
    AddM "x5 Вклад спроса", Contrib("DIVIDE([x5_w TOTAL_Шт], [x5_w TDP])", "DIVIDE([x5_c TOTAL_Шт], [x5_c TDP])"), fDec1
    AddM "x5 Вклад веса", Contrib("DIVIDE([x5_w TOTAL_КгЛ], [x5_w TOTAL_Шт])", "DIVIDE([x5_c TOTAL_КгЛ], [x5_c TOTAL_Шт])"), fDec1
    AddM "x5 Вклад цены", Contrib("DIVIDE([x5_w TOTAL_Руб], [x5_w TOTAL_КгЛ])", "DIVIDE([x5_c TOTAL_Руб], [x5_c TOTAL_КгЛ])"), fDec1
    AddM "x5 Изм продаж руб", "VAR vR1 = [x5_w TOTAL_Руб] VAR vR0 = [x5_c TOTAL_Руб] " & _
                              "RETURN IF(vR0 > 0, (vR1 - vR0) / 1000000)", fDec1

    ' --- Помесячные продажи для графиков ----------------------------------------------
    AddM "x5 Продажи мес", "DIVIDE(SWITCH([x5_Единица], ""кг"", SUM('Факт'[TOTAL_КгЛ]), ""шт"", SUM('Факт'[TOTAL_Шт]), " & _
                           "SUM('Факт'[TOTAL_Руб])), [x5_Масштаб])", fDec1
    AddM "x5 Продажи YA мес", "VAR vE = MAX('Календарь'[НачалоПериода]) VAR vP = EDATE(vE, -12) " & _
                              "RETURN CALCULATE([x5 Продажи мес], FILTER(ALL('Календарь'), 'Календарь'[НачалоПериода] = vP))", fDec1

    WriteLog "Меры созданы. Всего мер в модели: " & wb.Model.ModelMeasures.Count
End Sub

Private Sub AddM(ByVal nm As String, ByVal dax As String, ByVal fmt As Object)
    Dim d As String
    If known.Exists(nm) Then Err.Raise vbObjectError + 610, , "в модели уже есть мера «" & nm & "», созданная не сборщиком"
    On Error GoTo BadMeasure
    wb.Model.ModelMeasures.Add nm, wb.Model.ModelTables(T_REPORT), LocalDax(dax, ListSep()), fmt
    On Error GoTo 0
    known(nm) = True
    Status "Создаю меры: " & known.Count
    Exit Sub
BadMeasure:
    d = Err.Description
    WriteLog "Мера «" & nm & "» не создалась: " & d
    WriteLog "Её формула: " & LocalDax(dax, ListSep())
    Err.Raise vbObjectError + 611, , "мера «" & nm & "» не создалась, текст ошибки и формула на листе «" & SH_LOG & "»"
End Sub

' Разделитель аргументов в формулах Excel и Power Pivot на этом компьютере (в русской версии «;»)
Private Function ListSep() As String
    ListSep = CStr(Application.International(5))   ' xlListSeparator
    If Len(ListSep) = 0 Then ListSep = ","
End Function

' Меры записаны с запятой между аргументами. Power Pivot принимает формулу в местном виде,
' поэтому запятые вне строк в кавычках и вне имён в [ ] и ' ' заменяются на местный разделитель.
' Дробных чисел в мерах нет, поэтому десятичный разделитель не трогается.
Private Function LocalDax(ByVal dax As String, ByVal sep As String) As String
    Dim i As Long, ch As String, inText As Boolean, inBr As Boolean, inName As Boolean, out As String
    If sep = "," Then LocalDax = dax: Exit Function
    For i = 1 To Len(dax)
        ch = Mid$(dax, i, 1)
        If ch = Chr(34) And Not inBr And Not inName Then
            inText = Not inText
        ElseIf ch = "[" And Not inText And Not inName Then
            inBr = True
        ElseIf ch = "]" And inBr Then
            inBr = False
        ElseIf ch = "'" And Not inText And Not inBr Then
            inName = Not inName
        ElseIf ch = "," And Not inText And Not inBr And Not inName Then
            ch = sep
        End If
        out = out & ch
    Next
    LocalDax = out
End Function

Private Function VarsWin() As String
    VarsWin = "VAR vEnd = [x5_Конец] VAR vStart = [x5_Начало] "
End Function

Private Function VarsCmp() As String
    VarsCmp = "VAR vEnd = [x5_Конец] VAR vStart = [x5_Начало] VAR vShift = [x5_Сдвиг] " & _
              "VAR vS2 = EDATE(vStart, -vShift) VAR vE2 = EDATE(vEnd, -vShift) "
End Function

Private Function RangeWin() As String
    RangeWin = "FILTER(ALL('Календарь'), 'Календарь'[НачалоПериода] >= vStart && 'Календарь'[НачалоПериода] <= vEnd)"
End Function

Private Function RangeCmp() As String
    RangeCmp = "FILTER(ALL('Календарь'), 'Календарь'[НачалоПериода] >= vS2 && 'Календарь'[НачалоПериода] <= vE2)"
End Function

Private Function Win(ByVal inner As String) As String
    Win = VarsWin() & "RETURN CALCULATE(" & inner & ", " & RangeWin() & ")"
End Function

Private Function Cmp(ByVal inner As String) As String
    Cmp = VarsCmp() & "RETURN CALCULATE(" & inner & ", " & RangeCmp() & ")"
End Function

' Сумма в выбранной единице: prefix «x5_w TOTAL» -> [x5_w TOTAL_Руб] / _КгЛ / _Шт
Private Function UnitPick(ByVal prefix As String) As String
    UnitPick = "SWITCH([x5_Единица], ""кг"", [" & prefix & "_КгЛ], ""шт"", [" & prefix & "_Шт], [" & prefix & "_Руб])"
End Function

Private Function PriceIndex(ByVal p As String) As String
    Dim price As String
    price = "IF(vKg, DIVIDE([" & p & " TOTAL_Руб], [" & p & " TOTAL_КгЛ]), DIVIDE([" & p & " TOTAL_Руб], [" & p & " TOTAL_Шт]))"
    PriceIndex = "VAR vKg = [x5_Единица] = ""кг"" VAR vP = " & price & _
                 " VAR vC = CALCULATE(" & price & ", ALLEXCEPT('Продукты', 'Продукты'[Категория])) " & _
                 "RETURN IF(vP > 0 && vC > 0, vP / vC * 100)"
End Function

Private Function Contrib(ByVal nowExpr As String, ByVal cmpExpr As String) As String
    Contrib = "VAR vA = " & nowExpr & " VAR vB = " & cmpExpr & " VAR vL = [x5_L] " & _
              "RETURN IF(vA > 0 && vB > 0 && NOT ISBLANK(vL), vL * LN(vA / vB) / 1000000)"
End Function

'==============================================================================
'  Листы отчёта
'==============================================================================

Private Sub BuildSummary()
    Dim ws As Object, pt As Object, ch As Object, th As Variant, i As Long, r As Long
    Set ws = NewSheet(SH_SUM)
    SetGrid ws

    ' Сводная по производителям в скрытых столбцах. От неё создаются общие срезы книги.
    Set pt = NewPivot(ws, "AA6", "x5_Производители")
    pt.ManualUpdate = True
    AddRowField pt, "[Продукты].[" & COL_PROD & "]"
    AddDataField pt, "x5 Изм доли кат", "Изм. доли, п.п. "
    AddDataField pt, "x5 Доля кат", "Доля в категории "
    AddDataField pt, "x5 Продажи", "Продажи "
    StylePivot pt
    pt.ColumnGrand = False
    pt.RowGrand = False
    pt.ManualUpdate = False
    MakeCaches pt

    Header ws, "Что произошло за период и что это значит для Маревен"
    AddSlicerBand ws, "sum"

    Tile ws, 14, 2, 1, "=""Продажи категории, ""&" & CV("x5_Подпись единицы", "srUnit"), "x5 Продажи", F_NUM1, "x5 Прирост", F_UP_PCT
    Tile ws, 14, 3, 3, "=""Продажи Маревен, ""&" & CV("x5_Подпись единицы", "srUnit"), "x5 Продажи Маревен", F_NUM1, _
         "x5 Прирост Маревен", F_UP_PCT
    Tile ws, 14, 6, 3, "Доля Маревен в категории", "x5 Доля Маревен", "0.0%", "x5 Изм доли Маревен", F_UP_PP
    Tile ws, 14, 9, 3, "Индекс цены Маревен, категория = 100", "x5 Индекс цены Маревен", "0", "x5 Изм индекса Маревен", F_PTS
    Tile ws, 18, 2, 1, "Цена категории, руб/кг", "x5 Цена кг", F_NUM1, "x5 Изм цены кг", F_PCT
    Tile ws, 18, 3, 3, "Доля промо в категории", "x5 Доля промо", "0.0%", "x5 Изм доли промо", F_PP
    Tile ws, 18, 6, 3, "Доля промо у Маревен", "x5 Доля промо Маревен", "0.0%", "x5 Изм доли промо Маревен", F_PP
    Tile ws, 18, 9, 3, "Доля продаж по картам", "x5 Доля карты", "0.0%", "x5 Изм доли карты", F_PP

    ' Пять выигравших и пять проигравших по изменению доли (из сводной в AA:AD)
    ws.Range("AF7:AF506").Formula = "=IF(AND(AA7<>"""",ISNUMBER(AB7)),AB7+ROW()/1000000000,"""")"
    ws.Range("AH3").Formula = "=COUNT(AF7:AF506)"
    For i = 1 To 10
        ws.Cells(6 + i, 33).Value = i
    Next
    ws.Range("AH7:AH16").Formula = "=IF($AH$3<=10,IF(AG7<=$AH$3,AG7,""""),IF(AG7<=5,AG7,$AH$3-10+AG7))"
    ws.Range("AI7:AI16").Formula = "=IF(AH7="""","""",LARGE($AF$7:$AF$506,AH7))"
    ws.Range("AJ7:AJ16").Formula = "=IF(AH7="""","""",INDEX($AA$7:$AA$506,MATCH(AI7,$AF$7:$AF$506,0)))"
    ws.Range("AK7:AK16").Formula = "=IF(AJ7=""" & OWN & """,ROUND(AI7,3),"""")"
    ws.Range("AL7:AL16").Formula = "=IF(OR(AJ7="""",AJ7=""" & OWN & """),"""",ROUND(AI7,3))"

    SectionTitle ws, "B22", "Кто выиграл и кто проиграл долю, п.п."
    Set ch = NewChart(ws, "B23:F38")
    AddSeries ch, "Маревен", ws.Range("AK7:AK16"), ws.Range("AJ7:AJ16")
    AddSeries ch, "Другие производители", ws.Range("AL7:AL16"), ws.Range("AJ7:AJ16")
    StyleBars ch, Array(C_ACCENT, C_OTHER), "+0.0;-0.0;"

    SectionTitle ws, "G22", "Главное за период"
    th = Array("Рынок: размер, прирост и что его дало — объём или цена.", _
               "Где: сегмент, сеть или ФО с наибольшим вкладом в рост или падение.", _
               "Конкуренты: кто забрал долю и за счёт чего.", _
               "Маревен: доля, её изменение и главный драйвер.", _
               "Риск или возможность: цифра и следующий шаг.")
    For i = 0 To 4
        r = 23 + i * 3
        ws.Cells(r, 7).Value = (i + 1) & ". " & th(i)
        With ws.Range(ws.Cells(r, 7), ws.Cells(r + 1, 11))
            .Merge
            .WrapText = True
            .VerticalAlignment = XL_TOP
            .Font.Italic = True
            .Font.Color = C_MUTED
            .Interior.Color = C_TILE
            .IndentLevel = 1
        End With
    Next
    ws.Columns("AA:AL").Hidden = True
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    WriteLog "Лист «" & SH_SUM & "» построен"
End Sub

Private Sub BuildCategories()
    Dim ws As Object, pt As Object, ch As Object, i As Long, r As Long, labels As Variant, ms As Variant
    Set ws = NewSheet(SH_CAT)
    SetGrid ws
    Header ws, "Сколько стоит рынок, как он растёт и за счёт чего"
    AddSlicerBand ws, "cat"

    Tile ws, 14, 2, 1, "=""Продажи, ""&" & CV("x5_Подпись единицы", "srUnit"), "x5 Продажи", F_NUM1, "x5 Прирост", F_UP_PCT
    Tile ws, 14, 3, 3, "Цена, руб/кг", "x5 Цена кг", F_NUM1, "x5 Изм цены кг", F_PCT
    Tile ws, 14, 6, 3, "Цена, руб/шт", "x5 Цена шт", F_NUM1, "x5 Изм цены шт", F_PCT
    Tile ws, 14, 9, 3, "SKU с продажами", "x5 SKU", "#,##0", "x5 Изм SKU", F_INT
    Tile ws, 18, 2, 1, "TDP, среднее за окно", "x5 TDP", F_NUM1, "x5 Изм TDP", F_UP_PCT
    Tile ws, 18, 3, 3, "Доля промо", "x5 Доля промо", "0.0%", "x5 Изм доли промо", F_PP
    Tile ws, 18, 6, 3, "Доля продаж по картам", "x5 Доля карты", "0.0%", "x5 Изм доли карты", F_PP
    Tile ws, 18, 9, 3, "Доля Маревен", "x5 Доля Маревен", "0.0%", "x5 Изм доли Маревен", F_UP_PP

    ' Помесячный ряд за 25 месяцев до выбранного и год назад (скрытые столбцы AA:AH)
    ws.Range("AA3").Formula = "=IFERROR(" & CV("x5_Конец ГГГГММ", "srPeriod") & ",0)"
    For i = 0 To 24
        ws.Cells(5 + i, 27).Value = 24 - i
    Next
    ws.Range("AB5:AB29").Formula = "=INT($AA$3/100)*12+MOD($AA$3,100)-1-AA5"
    ws.Range("AC5:AC29").Formula = "=""[Календарь].[Период].&[""&(INT(AB5/12)*100+MOD(AB5,12)+1)&""]"""
    ws.Range("AD5:AD29").Formula = "=CHOOSE(MOD(AB5,12)+1,""янв"",""фев"",""мар"",""апр"",""май"",""июн""," & _
                                   """июл"",""авг"",""сен"",""окт"",""ноя"",""дек"")&"" ""&RIGHT(INT(AB5/12),2)"
    ws.Range("AE5:AE29").Formula = "=IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[x5 Продажи мес]"",AC5" & _
                                   SRL(MKT_SR & ",srUnit") & "),"""")"
    ws.Range("AF5:AF29").Formula = "=IF(ISNUMBER(AE5),AE5,NA())"
    ws.Range("AG5:AG29").Formula = "=IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[x5 Продажи YA мес]"",AC5" & _
                                   SRL(MKT_SR & ",srUnit") & "),"""")"
    ws.Range("AH5:AH29").Formula = "=IF(ISNUMBER(AG5),AG5,NA())"

    SectionTitle ws, "B22", "=""Продажи по месяцам, ""&" & CV("x5_Подпись единицы", "srUnit")
    Set ch = NewChart(ws, "B23:F38")
    AddSeries ch, "Продажи", ws.Range("AF5:AF29"), ws.Range("AD5:AD29")
    AddSeries ch, "Год назад", ws.Range("AH5:AH29"), ws.Range("AD5:AD29")
    StyleLines ch

    ' Разложение изменения продаж в рублях (скрытые столбцы AA:AE, строки 33-37)
    labels = Array("Полка (TDP)", "Спрос на точку", "Вес упаковки", "Цена за кг", "Итого")
    ms = Array("x5 Вклад полки", "x5 Вклад спроса", "x5 Вклад веса", "x5 Вклад цены", "x5 Изм продаж руб")
    For i = 0 To 4
        r = 33 + i
        ws.Cells(r, 27).Value = labels(i)
        ws.Cells(r, 28).Formula = "=IFERROR(" & CV(CStr(ms(i)), ALL_SR) & ","""")"
        If i < 4 Then
            ws.Cells(r, 29).Formula = "=IF(AND(ISNUMBER(AB" & r & "),AB" & r & ">=0),AB" & r & ","""")"
            ws.Cells(r, 30).Formula = "=IF(AND(ISNUMBER(AB" & r & "),AB" & r & "<0),AB" & r & ","""")"
        Else
            ws.Cells(r, 31).Formula = "=IF(ISNUMBER(AB" & r & "),AB" & r & ","""")"
        End If
    Next
    SectionTitle ws, "G22", "Чем изменились продажи, млн руб"
    Set ch = NewChart(ws, "G23:K38")
    AddSeries ch, "Рост", ws.Range("AC33:AC37"), ws.Range("AA33:AA37")
    AddSeries ch, "Снижение", ws.Range("AD33:AD37"), ws.Range("AA33:AA37")
    AddSeries ch, "Итого", ws.Range("AE33:AE37"), ws.Range("AA33:AA37")
    StyleBars ch, Array(C_GOOD, C_BAD, C_INK2), "+#,##0.0;-#,##0.0;"
    ch.HasLegend = False

    ' Таблица категорий
    SectionTitle ws, "B40", "Категории: продажи, доли и цены за окно"
    ws.Range("B41").Value = "Доля в ПБП — от всей Продукции быстрого приготовления. Доля Маревен — внутри категории."
    ws.Range("B41").Font.Color = C_INK2
    Set pt = NewPivot(ws, "B43", "x5_Категории")
    pt.ManualUpdate = True
    AddRowField pt, "[Продукты].[Категория]"
    AddDataField pt, "x5 Продажи", "Продажи "
    AddDataField pt, "x5 Прирост", "Прирост, % ", F_PIV_UP_PCT
    AddDataField pt, "x5 Доля ПБП", "Доля в ПБП "
    AddDataField pt, "x5 Изм доли ПБП", "Изм. доли в ПБП, п.п. ", F_PIV_PP
    AddDataField pt, "x5 Доля Маревен", "Доля Маревен "
    AddDataField pt, "x5 Изм доли Маревен", "Изм. доли Маревен, п.п. ", F_PIV_UP
    AddDataField pt, "x5 Цена кг", "Цена, руб/кг "
    AddDataField pt, "x5 Изм цены кг", "Изм. цены за кг, % ", F_PIV_PCT
    AddDataField pt, "x5 Доля промо", "Доля промо "
    AddDataField pt, "x5 SKU", "SKU с продажами "
    StylePivot pt
    pt.ManualUpdate = False
    ConnectPivot pt
    TrySort pt, "[Продукты].[Категория].[Категория]", "[Measures].[x5 Продажи]"
    On Error Resume Next
    With pt.TableRange1.Rows(1)
        .WrapText = True
        .RowHeight = 42
        .VerticalAlignment = XL_VCENTER
    End With
    On Error GoTo 0
    ws.Columns("AA:AL").Hidden = True
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    WriteLog "Лист «" & SH_CAT & "» построен"
End Sub

Private Sub BuildCover()
    Dim ws As Object, how As Variant, items As Variant, i As Long, r As Long
    Set ws = NewSheet(SH_COVER)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 110
    With ws.Range("B2")
        .Value = "X5 Dialog · Продукция быстрого приготовления"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = C_INK
    End With
    ws.Range("Z2").Formula = "=IFERROR(" & CV("x5_Первый ГГГГММ", "") & ",0)"
    ws.Range("Z3").Formula = "=IFERROR(" & CV("x5_Последний ГГГГММ", "") & ",0)"
    ws.Range("B3").Formula = "=IFERROR(""Месячный отчёт · данные с ""&" & FxMonth("Z2") & "&"" по ""&" & FxMonth("Z3") & ","""")"
    ws.Range("B3").Font.Color = C_INK2
    ws.Range("B4").Value = "Собрано: " & Format(Now, "dd.mm.yyyy hh:nn")
    ws.Range("B4").Font.Color = C_MUTED

    SectionTitle ws, "B6", "Как пользоваться"
    how = Array("Срезы наверху листов общие для всей книги: сеть, выбранная на одном листе, выбрана на всех.", _
                "«Последний месяц»: без выбора отчёт строится на последнем месяце в данных.", _
                "«Окно»: Месяц, L3M, YTD или MAT, без выбора — MAT. «Сравнение»: YA или PP, без выбора — YA.", _
                "«Единица»: руб, кг или шт. Продажи — в млн руб, тоннах или млн шт; цены всегда в рублях.", _
                "Маревен выделена цветом и стоит на своём месте в рейтингах.", _
                "Обновление: Данные " & ARW & " Обновить всё (или макрос ОбновитьОтчёт), затем лист «" & SH_CHK & "».")
    For i = 0 To UBound(how)
        ws.Cells(7 + i, 2).Value = "• " & how(i)
    Next

    SectionTitle ws, "B14", "Листы"
    items = Array(Array(SH_SUM, "Что произошло за период и что это значит для Маревен"), _
                  Array(SH_CAT, "Сколько стоит рынок, как он растёт и за счёт чего"), _
                  Array(SH_CHK, "Сходятся ли меры отчёта с суммами по месяцам"))
    For i = 0 To UBound(items)
        r = 15 + i
        ws.Hyperlinks.Add ws.Cells(r, 2), "", "'" & items(i)(0) & "'!A1", , CStr(items(i)(0))
        ws.Cells(r, 3).Value = items(i)(1)
    Next
    ws.Cells(19, 2).Value = "Следующие листы: Сегменты, Сети и регионы, Производители, Бренды и линии, SKU, Цены и промо, Лояльные покупки."
    ws.Cells(19, 2).Font.Color = C_MUTED
    ws.Columns("Z").Hidden = True
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    WriteLog "Лист «" & SH_COVER & "» построен"
End Sub

Private Sub BuildChecks()
    Dim ws As Object, i As Long, bridge As String
    Set ws = NewSheet(SH_CHK)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 66
    ws.Columns("C:E").ColumnWidth = 18
    With ws.Range("B1")
        .Value = "Сходятся ли меры отчёта с суммами по месяцам"
        .Font.Size = 16
        .Font.Bold = True
    End With
    ws.Range("B2").Value = "Считается при текущих срезах: последний месяц, сеть, ФО, регион, категория. Рубли, допуск 0,5 руб."
    ws.Range("B2").Font.Color = C_INK2

    ' Помесячные суммы за 36 месяцев до выбранного (скрытые столбцы AA:AD)
    ws.Range("AA3").Formula = "=IFERROR(" & CV("x5_Конец ГГГГММ", "srPeriod") & ",0)"
    For i = 0 To 35
        ws.Cells(5 + i, 27).Value = i
    Next
    ws.Range("AB5:AB40").Formula = "=INT($AA$3/100)*12+MOD($AA$3,100)-1-AA5"
    ws.Range("AC5:AC40").Formula = "=""[Календарь].[Период].&[""&(INT(AB5/12)*100+MOD(AB5,12)+1)&""]"""
    ws.Range("AD5:AD40").Formula = "=N(IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[x5_Руб мес]"",AC5" & SRL(MKT_SR) & "),0))"

    ws.Range("B4").Value = "Проверка"
    ws.Range("C4").Value = "Мера отчёта"
    ws.Range("D4").Value = "Сумма по месяцам"
    ws.Range("E4").Value = "Статус"
    ws.Range("B4:E4").Font.Bold = True
    CheckRow ws, 5, "MAT: окно = 12 последних месяцев", WinCube("x5_w TOTAL_Руб", "MAT", ""), "=SUM(AD5:AD16)", 0.5
    CheckRow ws, 6, "MAT к YA: сравнение = те же 12 месяцев год назад", WinCube("x5_c TOTAL_Руб", "MAT", "YA"), _
             "=SUM(AD17:AD28)", 0.5
    CheckRow ws, 7, "L3M к PP: сравнение = 3 месяца перед окном", WinCube("x5_c TOTAL_Руб", "L3M", "PP"), "=SUM(AD8:AD10)", 0.5
    CheckRow ws, 8, "YTD: окно = месяцы с января", WinCube("x5_w TOTAL_Руб", "YTD", ""), _
             "=SUMPRODUCT((AA5:AA40<MOD($AA$3,100))*AD5:AD40)", 0.5
    CheckRow ws, 9, "Месяц: окно = выбранный месяц", WinCube("x5_w TOTAL_Руб", "Месяц", ""), "=AD5", 0.5
    bridge = "=N(IFERROR(" & CV("x5 Вклад полки", ALL_SR) & ",0))+N(IFERROR(" & CV("x5 Вклад спроса", ALL_SR) & ",0))" & _
             "+N(IFERROR(" & CV("x5 Вклад веса", ALL_SR) & ",0))+N(IFERROR(" & CV("x5 Вклад цены", ALL_SR) & ",0))"
    CheckRow ws, 10, "Разложение: полка + спрос + вес + цена = изменение продаж, млн руб", bridge, _
             "=N(IFERROR(" & CV("x5 Изм продаж руб", ALL_SR) & ",0))", 0.001
    On Error Resume Next
    With ws.Range("E5:E10").FormatConditions
        .Delete
        With .Add(XL_CELL_VALUE, XL_EQUAL, "=""OK""")
            .Font.Color = C_GOOD
            .Font.Bold = True
        End With
        With .Add(XL_CELL_VALUE, XL_NOT_EQUAL, "=""OK""")
            .Font.Color = C_BAD
            .Font.Bold = True
        End With
    End With
    On Error GoTo 0
    ws.Range("B12").Formula = "=""Итого: ""&COUNTIF(E5:E10,""OK"")&"" из 6 проверок OK"""
    ws.Range("B12").Font.Bold = True
    ws.Range("B14").Value = "Внешняя сверка: в HTML-дашборде с тем же месяцем, окном, сравнением и рынком продажи и прирост " & _
                            "должны совпасть с плиткой на листе «" & SH_SUM & "»."
    ws.Range("B14").Font.Color = C_INK2
    ws.Columns("AA:AD").Hidden = True
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    WriteLog "Лист «" & SH_CHK & "» построен"
End Sub

Private Sub CheckRow(ByVal ws As Object, ByVal r As Long, ByVal label As String, ByVal f1 As String, _
                     ByVal f2 As String, ByVal tol As Double)
    ws.Cells(r, 2).Value = label
    ws.Cells(r, 3).Formula = f1
    ws.Cells(r, 4).Formula = f2
    ws.Cells(r, 5).Formula = "=IF(ABS(C" & r & "-D" & r & ")<=" & Replace(CStr(tol), ",", ".") & ",""OK"",""РАСХОЖДЕНИЕ"")"
    ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).NumberFormat = IIf(tol < 0.01, "#,##0.000", "#,##0")
End Sub

' Сумма в рублях за окно или сравнение при заданном окне — для листа «Проверка»
Private Function WinCube(ByVal measure As String, ByVal winName As String, ByVal cmpName As String) As String
    Dim m As String
    m = ",""[Окно].[Окно].&[" & winName & "]"""
    If Len(cmpName) > 0 Then m = m & ",""[Сравнение].[Сравнение].&[" & cmpName & "]"""
    WinCube = "=N(IFERROR(CUBEVALUE(""" & CN & """,""[Measures].[" & measure & "]""" & SRL("srPeriod," & MKT_SR) & m & "),0))"
End Function

Private Sub OrderSheets()
    On Error Resume Next
    wb.Worksheets(SH_COVER).Move Before:=wb.Worksheets(1)
    wb.Worksheets(SH_SUM).Move After:=wb.Worksheets(SH_COVER)
    wb.Worksheets(SH_CAT).Move After:=wb.Worksheets(SH_SUM)
    wb.Worksheets(SH_CHK).Move After:=wb.Worksheets(wb.Worksheets.Count)
    wb.Worksheets(SH_LOG).Move After:=wb.Worksheets(wb.Worksheets.Count)
    If Err.Number <> 0 Then WriteLog "Порядок листов: " & Err.Description
    On Error GoTo 0
End Sub

'==============================================================================
'  Общие элементы листов
'==============================================================================

Private Function NewSheet(ByVal nm As String) As Object
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then Err.Raise vbObjectError + 630, , "в книге уже есть лист «" & nm & _
                                        "», созданный не сборщиком. Переименуйте его и запустите сборку снова"
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = nm
    ws.CustomProperties.Add MARK, "1"
    ws.Cells.Font.Size = 10
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    On Error GoTo 0
    Set NewSheet = ws
End Function

' Сетка листа: A — поле, B — широкий столбец (плитка 1 и подписи сводных), C:N — по 11
Private Sub SetGrid(ByVal ws As Object)
    Dim r As Long
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C:N").ColumnWidth = 11
    ws.Rows(1).RowHeight = 28
    ws.Rows(2).RowHeight = 16
    For r = 3 To 12
        ws.Rows(r).RowHeight = 15.6
    Next
    ws.Rows(13).RowHeight = 8
    ws.Rows(14).RowHeight = 16
    ws.Rows(15).RowHeight = 30
    ws.Rows(16).RowHeight = 16
    ws.Rows(17).RowHeight = 8
    ws.Rows(18).RowHeight = 16
    ws.Rows(19).RowHeight = 30
    ws.Rows(20).RowHeight = 16
    ws.Rows(21).RowHeight = 10
End Sub

' Заголовок листа (вопрос, на который он отвечает) и строка с выбранными окном, сравнением, месяцем, единицей
Private Sub Header(ByVal ws As Object, ByVal title As String)
    With ws.Range("B1")
        .Value = title
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = C_INK
    End With
    ws.Range("AA2").Formula = "=IFERROR(" & CV("x5_Конец ГГГГММ", "srPeriod") & ",0)"
    ws.Range("B2").Formula = "=IFERROR(" & CV("x5_Окно", "srWindow") & "&"" к ""&" & CV("x5_Сравнение", "srCompare") & _
                             "&"" · последний месяц ""&" & FxMonth("AA2") & "&"" · продажи в ""&" & _
                             CV("x5_Подпись единицы", "srUnit") & ","""")"
    ws.Range("B2").Font.Color = C_INK2
End Sub

Private Sub SectionTitle(ByVal ws As Object, ByVal addr As String, ByVal text As String)
    With ws.Range(addr)
        If Left$(text, 1) = "=" Then .Formula = text Else .Value = text
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = C_INK
    End With
End Sub

' Плитка: подпись, значение за окно и изменение к сравнению
Private Sub Tile(ByVal ws As Object, ByVal r As Long, ByVal c1 As Long, ByVal nCols As Long, ByVal label As String, _
                 ByVal measure As String, ByVal valueFmt As String, ByVal deltaMeasure As String, ByVal deltaFmt As String)
    Dim c2 As Long
    c2 = c1 + nCols - 1
    With ws.Cells(r, c1)
        If Left$(label, 1) = "=" Then .Formula = label Else .Value = label
        .Font.Size = 9
        .Font.Color = C_INK2
    End With
    With ws.Cells(r + 1, c1)
        .Formula = "=IFERROR(" & CV(measure, ALL_SR) & ",""—"")"
        .NumberFormat = valueFmt
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = C_INK
    End With
    With ws.Cells(r + 2, c1)
        .Formula = "=IFERROR(" & CV(deltaMeasure, ALL_SR) & ","""")"
        .NumberFormat = deltaFmt
        .Font.Size = 9
        .Font.Color = C_INK2
    End With
    With ws.Range(ws.Cells(r, c1), ws.Cells(r + 2, c2))
        .Interior.Color = C_TILE
        .HorizontalAlignment = XL_LEFT
        .IndentLevel = 1
    End With
    ws.Range(ws.Cells(r, c1), ws.Cells(r, c2)).Merge
    ws.Range(ws.Cells(r + 1, c1), ws.Cells(r + 1, c2)).Merge
    ws.Range(ws.Cells(r + 2, c1), ws.Cells(r + 2, c2)).Merge
    With ws.Range(ws.Cells(r, c1), ws.Cells(r + 2, c1)).Borders(XL_EDGE_LEFT)
        .LineStyle = XL_CONTINUOUS
        .Weight = XL_THICK
        .Color = RGB(255, 255, 255)
    End With
End Sub

' Формула CUBEVALUE для меры отчёта с перечнем срезов
Private Function CV(ByVal measure As String, ByVal slicers As String) As String
    CV = "CUBEVALUE(""" & CN & """,""[Measures].[" & measure & "]""" & SRL(slicers) & ")"
End Function

' Срезы для формул: только созданные. Возвращает ",srPeriod,srWindow" или пустую строку
Private Function SRL(ByVal names As String) As String
    Dim p As Variant, s As String
    For Each p In Split(names, ",")
        s = Trim$(CStr(p))
        If Len(s) > 0 Then
            If srOk.Exists(s) Then SRL = SRL & "," & s
        End If
    Next
End Function

' Подписи срезов, которые не создались, через запятую
Private Function MissingSlicers() As String
    Dim nm As Variant, s As String
    If srOk Is Nothing Then Exit Function
    For Each nm In Split(CACHES, " ")
        If Not srOk.Exists(CStr(nm)) Then s = s & ", " & SlicerCaption(CStr(nm))
    Next
    If Len(s) > 0 Then MissingSlicers = Mid$(s, 3)
End Function

Private Function SlicerCaption(ByVal nm As String) As String
    Select Case nm
        Case "srPeriod": SlicerCaption = "Последний месяц"
        Case "srWindow": SlicerCaption = "Окно"
        Case "srCompare": SlicerCaption = "Сравнение"
        Case "srUnit": SlicerCaption = "Единица"
        Case "srNet": SlicerCaption = "Сеть"
        Case "srFO": SlicerCaption = "ФО"
        Case "srRegion": SlicerCaption = "Регион"
        Case "srCat": SlicerCaption = "Категория"
        Case Else: SlicerCaption = nm
    End Select
End Function

' Подпись месяца по коду ГГГГММ из ячейки: «октябрь 2026»
Private Function FxMonth(ByVal ref As String) As String
    FxMonth = "CHOOSE(MOD(" & ref & ",100),""январь"",""февраль"",""март"",""апрель"",""май"",""июнь"",""июль""," & _
              """август"",""сентябрь"",""октябрь"",""ноябрь"",""декабрь"")&"" ""&INT(" & ref & "/100)"
End Function

'==============================================================================
'  Срезы
'==============================================================================

Private Sub MakeCaches(ByVal pt As Object)
    MakeCache pt, "srPeriod", "[Календарь].[МесяцМетка]", XL_SORT_DESC
    MakeCache pt, "srWindow", "[Окно].[Окно]", XL_SORT_SOURCE
    MakeCache pt, "srCompare", "[Сравнение].[Сравнение]", XL_SORT_SOURCE
    MakeCache pt, "srUnit", "[Единица].[Единица]", XL_SORT_SOURCE
    MakeCache pt, "srNet", "[Сети].[Сеть]", XL_SORT_SOURCE
    MakeCache pt, "srFO", "[Регионы].[ФО]", XL_SORT_SOURCE
    MakeCache pt, "srRegion", "[Регионы].[СубъектРФ]", XL_SORT_SOURCE
    MakeCache pt, "srCat", "[Продукты].[Категория]", XL_SORT_SOURCE
    If srOk.Count = 0 Then
        WriteLog "Срезы не создались. Отчёт строится без них, на значениях по умолчанию: " & _
                 "последний месяц, MAT к YA, рубли, весь рынок"
    Else
        WriteLog "Срезы созданы: " & srOk.Count & " из 8 (способ " & scWay & ")"
    End If
End Sub

' Срезы создаются через типы Excel (ранняя привязка). Через «As Object» Excel получает имя поля
' по ссылке, и SlicerCaches.Add2 отвечает «Invalid procedure call or argument».
' Если срез не создаётся ни одним способом, сборка идёт дальше без него, причина — в журнале.
Private Sub MakeCache(ByVal pt As Object, ByVal nm As String, ByVal fld As String, ByVal sortMode As Long)
    Dim sc As Excel.SlicerCache, xpt As Excel.PivotTable, p As Excel.PivotTable, d As String, way As Long, linked As Boolean, lv As Object
    If scWay > 0 Then Set sc = AddCacheWay(scWay, pt, fld, nm, d)
    Do While sc Is Nothing And way < 5
        way = way + 1
        If way <> scWay Then
            Set sc = AddCacheWay(way, pt, fld, nm, d)
            If Not sc Is Nothing Then scWay = way
        End If
    Loop
    If sc Is Nothing Then
        WriteLog "Срез «" & SlicerCaption(nm) & "» (" & fld & ") не создался:" & d
        Exit Sub
    End If
    On Error Resume Next
    If sc.Name <> nm Then sc.Name = nm
    On Error GoTo 0
    If sc.Name <> nm Then
        WriteLog "Срез «" & SlicerCaption(nm) & "» получил имя «" & sc.Name & "» вместо «" & nm & _
                 "», формулы на него не сошлются. Убираю его"
        On Error Resume Next
        sc.Delete
        On Error GoTo 0
        Exit Sub
    End If
    srOk(nm) = True
    Set xpt = pt
    On Error Resume Next
    For Each p In sc.PivotTables
        If p.Name = xpt.Name And p.Parent.Name = xpt.Parent.Name Then linked = True
    Next
    Err.Clear
    If Not linked Then
        sc.PivotTables.AddPivotTable xpt
        If Err.Number <> 0 Then WriteLog "Сводная «" & xpt.Name & "» не подключилась к срезу " & nm & ": " & Err.Description
        Err.Clear
    End If
    sc.SortItems = sortMode
    Set lv = sc.SlicerCacheLevels(1)
    lv.SortItems = sortMode
    On Error GoTo 0
End Sub

' Один способ создать кэш среза. При ошибке возвращает Nothing и дописывает причину в d.
Private Function AddCacheWay(ByVal way As Long, ByVal pt As Object, ByVal fld As String, ByVal nm As String, _
                             ByRef d As String) As Excel.SlicerCache
    Dim xb As Excel.Workbook, xpt As Excel.PivotTable, o As Object
    Set xb = wb
    On Error Resume Next
    Select Case way
        Case 1   ' от модели данных, как «Вставка > Срез > Модель данных»
            Set AddCacheWay = xb.SlicerCaches.Add2(xb.Connections(CN), fld, nm)
        Case 2   ' от сводной, как «Анализ сводной таблицы > Вставить срез»
            Set xpt = pt
            Set AddCacheWay = xb.SlicerCaches.Add2(xpt, fld, nm)
        Case 3   ' от имени подключения
            Set AddCacheWay = xb.SlicerCaches.Add2(CN, fld, nm)
        Case 4   ' поздняя привязка, аргументы по значению
            Set o = wb.SlicerCaches.Add2(wb.Connections(CN), CStr(fld), CStr(nm))
            Set AddCacheWay = o
        Case 5   ' старый метод Add (Excel 2010)
            Set AddCacheWay = xb.SlicerCaches.Add(xb.Connections(CN), fld, nm)
    End Select
    If Err.Number <> 0 Then d = d & " | способ " & way & ": " & Err.Description
    Err.Clear
    If AddCacheWay Is Nothing Then Set AddCacheWay = xb.SlicerCaches(nm)   ' вдруг создался, но Excel ответил ошибкой
    Err.Clear
    On Error GoTo 0
End Function

' Полоса срезов в строках 3-12: сверху окно, сравнение, единица, сеть; снизу месяц, ФО, регион, категория
Private Sub AddSlicerBand(ByVal ws As Object, ByVal key As String)
    Dim L As Double, T As Double
    L = ws.Range("B3").Left
    T = ws.Range("B3").Top + 2
    PutSlicer ws, key, "srWindow", "[Окно].[Окно].[Окно]", "Окно", T, L, 200, 46, 4
    PutSlicer ws, key, "srCompare", "[Сравнение].[Сравнение].[Сравнение]", "Сравнение", T, L + 206, 110, 46, 2
    PutSlicer ws, key, "srUnit", "[Единица].[Единица].[Единица]", "Единица", T, L + 322, 150, 46, 3
    PutSlicer ws, key, "srNet", "[Сети].[Сеть].[Сеть]", "Сеть", T, L + 478, 247, 46, 3
    PutSlicer ws, key, "srPeriod", "[Календарь].[МесяцМетка].[МесяцМетка]", "Последний месяц", T + 52, L, 110, 100, 1
    PutSlicer ws, key, "srFO", "[Регионы].[ФО].[ФО]", "ФО", T + 52, L + 116, 170, 100, 1
    PutSlicer ws, key, "srRegion", "[Регионы].[СубъектРФ].[СубъектРФ]", "Регион", T + 52, L + 292, 200, 100, 1
    PutSlicer ws, key, "srCat", "[Продукты].[Категория].[Категория]", "Категория", T + 52, L + 498, 227, 100, 1
End Sub

Private Sub PutSlicer(ByVal ws As Object, ByVal key As String, ByVal cacheName As String, ByVal level As String, _
                      ByVal caption As String, ByVal T As Double, ByVal L As Double, ByVal W As Double, _
                      ByVal H As Double, ByVal nCols As Long)
    ' Тоже через типы Excel. Если срез не встал, сборка идёт дальше: формулы работают и без него на листе.
    Dim xb As Excel.Workbook, xws As Excel.Worksheet, sc As Excel.SlicerCache, sl As Excel.Slicer, sName As String, d As String
    If Not srOk.Exists(cacheName) Then Exit Sub
    Set xb = wb
    Set xws = ws
    sName = cacheName & "_" & key
    On Error Resume Next
    Set sc = xb.SlicerCaches(cacheName)
    If Not sc Is Nothing Then
        Set sl = sc.Slicers.Add(xws, level, sName, caption, T, L, W, H)
        If sl Is Nothing Then
            d = Err.Description
            Err.Clear
            Set sl = sc.Slicers.Add(xws, , sName, caption, T, L, W, H)
        End If
    End If
    If sl Is Nothing Then
        WriteLog "Срез «" & caption & "» не встал на лист «" & xws.Name & "»: " & d & " / " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    sl.Style = "SlicerStyleLight1"
    sl.NumberOfColumns = nCols
    sl.RowHeight = 16
    On Error GoTo 0
End Sub

' Подключает сводную ко всем созданным срезам книги
Private Sub ConnectPivot(ByVal pt As Object)
    Dim nm As Variant, xb As Excel.Workbook, xpt As Excel.PivotTable
    Set xb = wb
    Set xpt = pt
    For Each nm In Split(CACHES, " ")
        If srOk.Exists(CStr(nm)) Then
            On Error Resume Next
            xb.SlicerCaches(CStr(nm)).PivotTables.AddPivotTable xpt
            If Err.Number <> 0 Then WriteLog "Сводная «" & xpt.Name & "» не подключилась к срезу " & nm & ": " & Err.Description
            On Error GoTo 0
        End If
    Next
End Sub

'==============================================================================
'  Сводные
'==============================================================================

Private Function NewPivot(ByVal ws As Object, ByVal addr As String, ByVal nm As String) As Object
    Dim pc As Object
    Set pc = wb.PivotCaches.Create(XL_EXTERNAL, wb.Connections(CN), XL_PT_V15)
    Set NewPivot = pc.CreatePivotTable(ws.Range(addr), nm, , XL_PT_V15)
End Function

Private Sub AddRowField(ByVal pt As Object, ByVal fld As String)
    Dim d As String
    On Error GoTo NoField
    With pt.CubeFields(fld)
        .Orientation = XL_ROW_FIELD
        .Position = 1
    End With
    Exit Sub
NoField:
    d = Err.Description
    Err.Raise vbObjectError + 640, , "поле " & fld & " не встало в сводную (" & d & "). Если столбец скрыт " & _
        "в Power Pivot, покажите его: правый клик по столбцу > Показать в клиентских средствах"
End Sub

Private Sub AddDataField(ByVal pt As Object, ByVal measure As String, ByVal caption As String, _
                         Optional ByVal fmt As String = "")
    Dim f As Object
    Set f = pt.AddDataField(pt.CubeFields("[Measures].[" & measure & "]"))
    On Error Resume Next
    f.Caption = caption
    If Err.Number <> 0 Then WriteLog "Подпись столбца «" & Trim$(caption) & "» не применилась: " & Err.Description
    Err.Clear
    If Len(fmt) > 0 Then f.NumberFormat = fmt
    If Err.Number <> 0 Then WriteLog "Формат столбца «" & Trim$(caption) & "» не применился: " & Err.Description
    On Error GoTo 0
End Sub

Private Sub StylePivot(ByVal pt As Object)
    On Error Resume Next
    pt.RowAxisLayout XL_TABULAR
    pt.TableStyle2 = "PivotStyleLight1"
    pt.ShowTableStyleRowStripes = False
    pt.HasAutoFormat = False
    pt.GrandTotalName = "Итого"
    On Error GoTo 0
End Sub

Private Sub TrySort(ByVal pt As Object, ByVal fld As String, ByVal measureName As String)
    On Error Resume Next
    pt.PivotFields(fld).AutoSort XL_DESCENDING, measureName
    If Err.Number <> 0 Then WriteLog "Сортировка сводной «" & pt.Name & "» не применилась: " & Err.Description
    On Error GoTo 0
End Sub

'==============================================================================
'  Графики
'==============================================================================

Private Function NewChart(ByVal ws As Object, ByVal area As String) As Object
    Dim rng As Object, co As Object
    Set rng = ws.Range(area)
    Set co = ws.ChartObjects.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    On Error Resume Next
    With co.Chart
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        .PlotVisibleOnly = False
        .HasTitle = False
        .ChartArea.Format.Line.Visible = MSO_FALSE
        .ChartArea.Font.Size = 9
        .ChartArea.Font.Color = C_INK2
    End With
    If Err.Number <> 0 Then WriteLog "Оформление графика: " & Err.Description
    On Error GoTo 0
    Set NewChart = co.Chart
End Function

Private Sub AddSeries(ByVal ch As Object, ByVal nm As String, ByVal vals As Object, ByVal cats As Object)
    ' Ряд ссылается на диапазон. Если Excel не примет диапазон, ряд задаётся формулой =SERIES(...).
    ' Ошибка графика не останавливает сборку, она пишется в журнал.
    Dim s As Object
    On Error Resume Next
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.Values = vals
    s.XValues = cats
    If Err.Number <> 0 Then
        Err.Clear
        s.Formula = "=SERIES(""" & nm & """," & RefA1(cats) & "," & RefA1(vals) & "," & ch.SeriesCollection.Count & ")"
        If Err.Number <> 0 Then WriteLog "Ряд «" & nm & "» графика не задался: " & Err.Description
    End If
    On Error GoTo 0
End Sub

Private Function RefA1(ByVal rng As Object) As String
    RefA1 = "'" & Replace(rng.Worksheet.Name, "'", "''") & "'!" & rng.Address
End Function

' Линии: текущий период цветом, год назад серым пунктиром
Private Sub StyleLines(ByVal ch As Object)
    ch.ChartType = XL_LINE
    On Error Resume Next
    With ch.SeriesCollection(1).Format.Line
        .ForeColor.RGB = C_ACCENT
        .Weight = 2.25
    End With
    With ch.SeriesCollection(2).Format.Line
        .ForeColor.RGB = C_MUTED
        .Weight = 1.5
        .DashStyle = MSO_DASH
    End With
    With ch.Axes(XL_VALUE)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = C_GRID
        .TickLabels.NumberFormat = "#,##0"
        .Format.Line.Visible = MSO_FALSE
    End With
    With ch.Axes(XL_CATEGORY)
        .TickLabelSpacing = 3
        .Format.Line.ForeColor.RGB = C_GRID
    End With
    ch.HasLegend = True
    ch.Legend.Position = XL_TOP
    If Err.Number <> 0 Then WriteLog "Оформление линейного графика: " & Err.Description
    On Error GoTo 0
End Sub

' Горизонтальные столбцы от нуля: ряды накладываются, пустые ячейки не рисуются
Private Sub StyleBars(ByVal ch As Object, ByVal colors As Variant, ByVal labelFmt As String)
    Dim i As Long, s As Object
    ch.ChartType = XL_BAR
    On Error Resume Next
    ch.ChartGroups(1).Overlap = 100
    ch.ChartGroups(1).GapWidth = 50
    For i = 1 To ch.SeriesCollection.Count
        Set s = ch.SeriesCollection(i)
        s.Format.Fill.ForeColor.RGB = colors(i - 1)
        s.HasDataLabels = True
        s.DataLabels.NumberFormat = labelFmt
        s.DataLabels.Position = XL_LABEL_OUT
        s.DataLabels.Font.Size = 8
    Next
    With ch.Axes(XL_CATEGORY)
        .ReversePlotOrder = True
        .TickLabelPosition = XL_TICK_LOW
        .Format.Line.Visible = MSO_FALSE
    End With
    With ch.Axes(XL_VALUE)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = C_GRID
        .TickLabels.NumberFormat = "+0.0;-0.0;0"
        .Format.Line.Visible = MSO_FALSE
    End With
    ch.HasLegend = True
    ch.Legend.Position = XL_TOP
    If Err.Number <> 0 Then WriteLog "Оформление столбчатого графика: " & Err.Description
    On Error GoTo 0
End Sub
