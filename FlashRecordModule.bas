Attribute VB_Name = "FlashRecordModule"
Option Explicit

Private Type SummaryStats
    FilesProcessed As Long
    RowsImported As Long
    Duplicates As Long
    Errors As Long
    SkippedFiles As Long
End Type

Private Const ROOT_PATH As String = "Z:\O Operations\O17 CCR files\O17.04 220 kV Sys\O17.04.01 Relay Room\Load profiles\Flash Record 2025"

Private gStats As SummaryStats
Private gFSO As Object
Private Const DICT_COMPARE_TEXT As Long = 1&

Public Sub BuildFlashRecordReadings()
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim dictReadings As Object
    Dim wsAll As Worksheet
    Dim wsDaily As Worksheet
    Dim wsErrors As Worksheet
    Dim wsSkipped As Worksheet
    Dim wsDuplicates As Worksheet
    Dim records As Variant
    Dim dailyRecords As Variant

    On Error GoTo CleanFail

    Set gFSO = CreateObject("Scripting.FileSystemObject")
    If Not gFSO.FolderExists(ROOT_PATH) Then
        Err.Raise vbObjectError + 1000, "BuildFlashRecordReadings", _
                  "Root folder not found: " & ROOT_PATH
    End If

    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ResetStats gStats

    Set dictReadings = CreateObject("Scripting.Dictionary")
    dictReadings.CompareMode = DICT_COMPARE_TEXT
    dictReadings.CompareMode = vbTextCompare

    Set wsAll = EnsureOutputSheet("All Readings", Array( _
        "Time Stamp", "Meter 1 Import MWh", "Meter 1 Export MWh", _
        "Meter 1 Source File", "Meter 1 Source Row", "Meter 1 Source Address", _
        "Meter 2 Import MWh", "Meter 2 Export MWh", "Meter 2 Source File", _
        "Meter 2 Source Address"))
    dictReadings.CompareMode = TextCompare

    Set wsAll = EnsureOutputSheet("All Readings", Array("Time Stamp", "Meter 1 Import MWh", "Meter 1 Export MWh", "Meter 2 Import MWh", "Meter 2 Export MWh"))
    Set wsDaily = EnsureOutputSheet("Daily Readings", Array("Date", "Meter 1 Import (Sum)", "Meter 1 Export (Sum)", "Meter 2 Import (Sum)", "Meter 2 Export (Sum)"))

    Set wsErrors = EnsureLogSheet("_Errors", Array("File Path", "Line", "Raw Time Stamp", "Issue"))
    Set wsSkipped = EnsureLogSheet("_SkippedFiles", Array("File Path", "Reason"))
    Set wsDuplicates = EnsureLogSheet("_Duplicates", Array("File Path", "Time Stamp", "Meter", "Action"))

    ProcessFolder ROOT_PATH, dictReadings, wsErrors, wsSkipped, wsDuplicates

    records = BuildSortedRecordArray(dictReadings)
    PopulateAllReadings wsAll, records

    dailyRecords = BuildDailyRecords(records)
    PopulateDailyReadings wsDaily, dailyRecords

    HideWorksheet wsErrors
    HideWorksheet wsSkipped
    HideWorksheet wsDuplicates

    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEnableEvents

    ShowSummaryMessage

    Exit Sub

CleanFail:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEnableEvents

    gStats.Errors = gStats.Errors + 1
    MsgBox "Flash Record build failed: " & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub ProcessFolder(ByVal folderPath As String, ByVal dictReadings As Object, _
                          ByVal wsErrors As Worksheet, ByVal wsSkipped As Worksheet, _
                          ByVal wsDuplicates As Worksheet)
    Dim folder As Object
    Dim subFolder As Object
    Dim fileItem As Object

    Set folder = gFSO.GetFolder(folderPath)

    For Each fileItem In folder.Files
        If LCase$(gFSO.GetExtensionName(fileItem.Path)) = "xls" Then
            ProcessDataFile fileItem.Path, dictReadings, wsErrors, wsSkipped, wsDuplicates
        End If
    Next fileItem

    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder.Path, dictReadings, wsErrors, wsSkipped, wsDuplicates
    Next subFolder
End Sub

Private Sub ProcessDataFile(ByVal filePath As String, ByVal dictReadings As Object, _
                            ByVal wsErrors As Worksheet, ByVal wsSkipped As Worksheet, _
                            ByVal wsDuplicates As Worksheet)
    Dim meterId As Long
    Dim fileNum As Integer
    Dim lineText As String
    Dim lineNo As Long
    Dim fields As Variant
    Dim tsValue As Double
    Dim importValue As Double
    Dim exportValue As Double
    Dim key As String
    Dim entry As Variant
    Dim meterLabel As String

    meterId = DetermineMeterId(gFSO.GetFileName(filePath))
    If meterId = 0 Then
        LogSkipped wsSkipped, filePath, "Meter identifier not found"
        gStats.SkippedFiles = gStats.SkippedFiles + 1
        Exit Sub
    End If

    gStats.FilesProcessed = gStats.FilesProcessed + 1
    meterLabel = "Meter " & CStr(meterId)

    On Error GoTo FileOpenFail
    fileNum = FreeFile
    Open filePath For Input As #fileNum

    On Error GoTo FileReadFail
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineNo = lineNo + 1

        If Trim$(lineText) = vbNullString Then
            GoTo ContinueLoop
        End If

        fields = Split(lineText, vbTab)
        If UBound(fields) < 3 Then
            LogError wsErrors, filePath, lineNo, GetField(fields, 0), "Insufficient columns"
            gStats.Errors = gStats.Errors + 1
            GoTo ContinueLoop
        End If

        If Not TryParseDateTime(GetField(fields, 0), tsValue) Then
            LogError wsErrors, filePath, lineNo, GetField(fields, 0), "Unrecognized date/time format"
            gStats.Errors = gStats.Errors + 1
            GoTo ContinueLoop
        End If

        If Not TryParseDouble(GetField(fields, 2), importValue) Then
            LogError wsErrors, filePath, lineNo, GetField(fields, 0), "Invalid import value"
            gStats.Errors = gStats.Errors + 1
            GoTo ContinueLoop
        End If

        If Not TryParseDouble(GetField(fields, 3), exportValue) Then
            LogError wsErrors, filePath, lineNo, GetField(fields, 0), "Invalid export value"
            gStats.Errors = gStats.Errors + 1
            GoTo ContinueLoop
        End If

        key = Format$(tsValue, "yyyy-mm-dd hh:nn:ss")
        If dictReadings.Exists(key) Then
            entry = dictReadings(key)
        Else
            ReDim entry(1 To 9)
            ReDim entry(1 To 5)
            entry(1) = tsValue
            entry(2) = Empty
            entry(3) = Empty
            entry(4) = Empty
            entry(5) = Empty
            entry(6) = vbNullString
            entry(7) = 0
            entry(8) = vbNullString
            entry(9) = 0
        End If

        If meterId = 1 Then
            If Not IsEmpty(entry(2)) Or Not IsEmpty(entry(3)) Then
                LogDuplicate wsDuplicates, filePath, tsValue, meterLabel, "Replaced existing values"
                gStats.Duplicates = gStats.Duplicates + 1
            End If
            entry(2) = importValue
            entry(3) = exportValue
            entry(6) = filePath
            entry(7) = lineNo
        Else
            If Not IsEmpty(entry(4)) Or Not IsEmpty(entry(5)) Then
                LogDuplicate wsDuplicates, filePath, tsValue, meterLabel, "Replaced existing values"
                gStats.Duplicates = gStats.Duplicates + 1
            End If
            entry(4) = importValue
            entry(5) = exportValue
            entry(8) = filePath
            entry(9) = lineNo
        End If

        dictReadings(key) = entry
        gStats.RowsImported = gStats.RowsImported + 1

ContinueLoop:
    Loop

    Close #fileNum
    Exit Sub

FileOpenFail:
    LogError wsErrors, filePath, 0, vbNullString, "Unable to open file"
    gStats.Errors = gStats.Errors + 1
    Exit Sub

FileReadFail:
    Close #fileNum
    LogError wsErrors, filePath, lineNo, vbNullString, "Unexpected read error: " & Err.Description
    gStats.Errors = gStats.Errors + 1
End Sub

Private Function BuildSortedRecordArray(ByVal dictReadings As Object) As Variant
    Dim records As Variant
    Dim i As Long
    Dim temp() As Variant

    If dictReadings Is Nothing Or dictReadings.Count = 0 Then
        BuildSortedRecordArray = Array()
        Exit Function
    End If

    records = dictReadings.Items
    ' Convert to 1-based array for sorting
    ReDim temp(1 To UBound(records) - LBound(records) + 1)
    For i = LBound(records) To UBound(records)
        temp(i - LBound(records) + 1) = records(i)
    Next i
    SortRecordArray temp, LBound(temp), UBound(temp)
    BuildSortedRecordArray = temp
End Function

Private Sub SortRecordArray(ByRef arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim low As Long
    Dim high As Long
    Dim mid As Long
    Dim pivot As Double
    Dim temp As Variant

    If first >= last Then Exit Sub

    mid = (first + last) \ 2
    pivot = arr(mid)(1)
    low = first
    high = last

    Do While low <= high
        Do While arr(low)(1) < pivot
            low = low + 1
        Loop
        Do While arr(high)(1) > pivot
            high = high - 1
        Loop
        If low <= high Then
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then SortRecordArray arr, first, high
    If low < last Then SortRecordArray arr, low, last
End Sub

Private Function BuildDailyRecords(ByVal records As Variant) As Variant
    Dim dictDaily As Object
    Dim i As Long
    Dim entry As Variant
    Dim key As String
    Dim dayEntry As Variant
    Dim items As Variant
    Dim temp() As Variant
    Dim lower As Long
    Dim upper As Long

    If GetArraySize(records) = 0 Then
        BuildDailyRecords = Array()
        Exit Function
    End If

    Set dictDaily = CreateObject("Scripting.Dictionary")

    lower = LBound(records)
    upper = UBound(records)

    For i = lower To upper
        entry = records(i)
        key = Format$(Int(entry(1)), "yyyy-mm-dd")
        If dictDaily.Exists(key) Then
            dayEntry = dictDaily(key)
        Else
            ReDim dayEntry(1 To 5)
            dayEntry(1) = Int(entry(1))
            dayEntry(2) = 0#
            dayEntry(3) = 0#
            dayEntry(4) = 0#
            dayEntry(5) = 0#
        End If

        dayEntry(2) = dayEntry(2) + GetNumericValue(entry(2))
        dayEntry(3) = dayEntry(3) + GetNumericValue(entry(3))
        dayEntry(4) = dayEntry(4) + GetNumericValue(entry(4))
        dayEntry(5) = dayEntry(5) + GetNumericValue(entry(5))

        dictDaily(key) = dayEntry
    Next i

    If dictDaily.Count = 0 Then
        BuildDailyRecords = Array()
        Exit Function
    End If

    items = dictDaily.Items
    ReDim temp(1 To UBound(items) - LBound(items) + 1)
    For i = LBound(items) To UBound(items)
        temp(i - LBound(items) + 1) = items(i)
    Next i

    SortRecordArray temp, LBound(temp), UBound(temp)
    BuildDailyRecords = temp
End Function

Private Sub PopulateAllReadings(ByVal ws As Worksheet, ByVal records As Variant)
    Dim dataArr() As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim tblRange As Range
    Dim tbl As ListObject
    Dim prevSheet As Worksheet
    Dim lower As Long
    Dim upper As Long

    ws.Cells.Clear
    WriteHeaders ws, Array( _
        "Time Stamp", "Meter 1 Import MWh", "Meter 1 Export MWh", _
        "Meter 1 Source File", "Meter 1 Source Row", "Meter 1 Source Address", _
        "Meter 2 Import MWh", "Meter 2 Export MWh", "Meter 2 Source File", _
        "Meter 2 Source Address")
    WriteHeaders ws, Array("Time Stamp", "Meter 1 Import MWh", "Meter 1 Export MWh", "Meter 2 Import MWh", "Meter 2 Export MWh")

    rowCount = GetArraySize(records)
    If rowCount > 0 Then
        lower = LBound(records)
        upper = UBound(records)
        ReDim dataArr(1 To rowCount, 1 To 10)
        ReDim dataArr(1 To rowCount, 1 To 5)
        For i = lower To upper
            dataArr(i - lower + 1, 1) = records(i)(1)
            dataArr(i - lower + 1, 2) = ToDisplayValue(records(i)(2))
            dataArr(i - lower + 1, 3) = ToDisplayValue(records(i)(3))
            dataArr(i - lower + 1, 4) = ToDisplayValue(records(i)(6))
            dataArr(i - lower + 1, 5) = ToSourceRowValue(records(i)(7))
            dataArr(i - lower + 1, 6) = FormatSourceAddress(records(i)(6), records(i)(7))
            dataArr(i - lower + 1, 7) = ToDisplayValue(records(i)(4))
            dataArr(i - lower + 1, 8) = ToDisplayValue(records(i)(5))
            dataArr(i - lower + 1, 9) = ToDisplayValue(records(i)(8))
            dataArr(i - lower + 1, 10) = FormatSourceAddress(records(i)(8), records(i)(9))
        Next i
        ws.Range("A2").Resize(rowCount, 10).Value = dataArr
        Set tblRange = ws.Range("A1").Resize(rowCount + 1, 10)
            dataArr(i - lower + 1, 4) = ToDisplayValue(records(i)(4))
            dataArr(i - lower + 1, 5) = ToDisplayValue(records(i)(5))
        Next i
        ws.Range("A2").Resize(rowCount, 5).Value = dataArr
        Set tblRange = ws.Range("A1").Resize(rowCount + 1, 5)
        On Error Resume Next
        ws.ListObjects("tblAllReadings").Delete
        On Error GoTo 0
        Set tbl = ws.ListObjects.Add(xlSrcRange:=tblRange, XlListObjectHasHeaders:=xlYes)
        tbl.Name = "tblAllReadings"
        tbl.TableStyle = "TableStyleMedium2"
    Else
        On Error Resume Next
        ws.ListObjects("tblAllReadings").Delete
        On Error GoTo 0
    End If

    ws.Columns("A:J").AutoFit
    ws.Columns("A").NumberFormat = "dd.mmm.yyyy hh:mm"
    ws.Columns("B:C").NumberFormat = "#,##0"
    ws.Columns("G:H").NumberFormat = "#,##0"
    ws.Columns("E").NumberFormat = "0"
    ws.Columns("D:F").NumberFormat = "@"
    ws.Columns("I:J").NumberFormat = "@"
    ws.Columns("A:E").AutoFit
    ws.Columns("A").NumberFormat = "dd.mmm.yyyy hh:mm"
    ws.Columns("B:E").NumberFormat = "#,##0"
    ws.Rows(1).Font.Bold = True

    Set prevSheet = Nothing
    On Error Resume Next
    Set prevSheet = ThisWorkbook.ActiveSheet
    On Error GoTo 0

    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With

    If Not prevSheet Is Nothing Then
        prevSheet.Activate
    End If
End Sub

Private Sub PopulateDailyReadings(ByVal ws As Worksheet, ByVal records As Variant)
    Dim dataArr() As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim tblRange As Range
    Dim tbl As ListObject
    Dim prevSheet As Worksheet
    Dim lower As Long
    Dim upper As Long

    ws.Cells.Clear
    WriteHeaders ws, Array("Date", "Meter 1 Import (Sum)", "Meter 1 Export (Sum)", "Meter 2 Import (Sum)", "Meter 2 Export (Sum)")

    rowCount = GetArraySize(records)
    If rowCount > 0 Then
        lower = LBound(records)
        upper = UBound(records)
        ReDim dataArr(1 To rowCount, 1 To 5)
        For i = lower To upper
            dataArr(i - lower + 1, 1) = records(i)(1)
            dataArr(i - lower + 1, 2) = records(i)(2)
            dataArr(i - lower + 1, 3) = records(i)(3)
            dataArr(i - lower + 1, 4) = records(i)(4)
            dataArr(i - lower + 1, 5) = records(i)(5)
        Next i
        ws.Range("A2").Resize(rowCount, 5).Value = dataArr
        Set tblRange = ws.Range("A1").Resize(rowCount + 1, 5)
        On Error Resume Next
        ws.ListObjects("tblDailyReadings").Delete
        On Error GoTo 0
        Set tbl = ws.ListObjects.Add(xlSrcRange:=tblRange, XlListObjectHasHeaders:=xlYes)
        tbl.Name = "tblDailyReadings"
        tbl.TableStyle = "TableStyleMedium2"
    Else
        On Error Resume Next
        ws.ListObjects("tblDailyReadings").Delete
        On Error GoTo 0
    End If

    ws.Columns("A:E").AutoFit
    ws.Columns("A").NumberFormat = "dd.mmm.yyyy"
    ws.Columns("B:E").NumberFormat = "#,##0"
    ws.Rows(1).Font.Bold = True

    Set prevSheet = Nothing
    On Error Resume Next
    Set prevSheet = ThisWorkbook.ActiveSheet
    On Error GoTo 0

    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With

    If Not prevSheet Is Nothing Then
        prevSheet.Activate
    End If
End Sub

Private Function GetArraySize(ByVal arr As Variant) As Long
    On Error GoTo Failed
    If IsArray(arr) Then
        GetArraySize = UBound(arr) - LBound(arr) + 1
    End If
    Exit Function
Failed:
    GetArraySize = 0
End Function

Private Function DetermineMeterId(ByVal fileName As String) As Long
    Dim lowerName As String
    lowerName = LCase$(fileName)

    If RegexTest(lowerName, "\bm(?:eter)?\s*(?:#|no\.?|number)?\s*0*1\b") Then
        DetermineMeterId = 1
    ElseIf RegexTest(lowerName, "\bm(?:eter)?\s*(?:#|no\.?|number)?\s*0*2\b") Then
        DetermineMeterId = 2
    Else
        DetermineMeterId = 0
    End If
End Function

Private Function RegexTest(ByVal text As String, ByVal pattern As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.Global = False
    re.IgnoreCase = True
    RegexTest = re.Test(text)
End Function

Private Function TryParseDateTime(ByVal value As String, ByRef result As Double) As Boolean
    Dim cleaned As String
    Dim alt As String

    TryParseDateTime = False
    cleaned = NormalizeDateTimeString(value)
    If cleaned = vbNullString Then Exit Function

    On Error GoTo FailParse
    If IsDate(cleaned) Then
        result = CDbl(CDate(cleaned))
        TryParseDateTime = True
        Exit Function
    End If

    alt = Replace(cleaned, ".", "/")
    If alt <> cleaned And IsDate(alt) Then
        result = CDbl(CDate(alt))
        TryParseDateTime = True
        Exit Function
    End If

    alt = Replace(cleaned, "-", "/")
    If alt <> cleaned And IsDate(alt) Then
        result = CDbl(CDate(alt))
        TryParseDateTime = True
        Exit Function
    End If

    alt = Replace(Replace(cleaned, ".", "/"), "-", "/")
    If alt <> cleaned And IsDate(alt) Then
        result = CDbl(CDate(alt))
        TryParseDateTime = True
        Exit Function
    End If

FailParse:
    TryParseDateTime = False
End Function

Private Function NormalizeDateTimeString(ByVal value As String) As String
    Dim cleaned As String

    cleaned = Trim$(value)
    cleaned = Replace(cleaned, Chr$(9), " ")
    Do While InStr(cleaned, "  ") > 0
        cleaned = Replace(cleaned, "  ", " ")
    Loop

    cleaned = Replace(cleaned, "/ ", "/")
    cleaned = Replace(cleaned, " /", "/")
    cleaned = Replace(cleaned, "- ", "-")
    cleaned = Replace(cleaned, " -", "-")
    cleaned = Replace(cleaned, ". ", ".")
    cleaned = Replace(cleaned, " .", ".")

    NormalizeDateTimeString = cleaned
End Function

Private Function TryParseDouble(ByVal value As String, ByRef result As Double) As Boolean
    Dim cleaned As String

    TryParseDouble = False
    cleaned = Trim$(value)
    cleaned = Replace(cleaned, ",", vbNullString)
    If cleaned = vbNullString Then
        result = 0#
        TryParseDouble = True
        Exit Function
    End If

    If IsNumeric(cleaned) Then
        result = CDbl(cleaned)
        TryParseDouble = True
    End If
End Function

Private Function GetNumericValue(ByVal value As Variant) As Double
    If IsError(value) Then
        GetNumericValue = 0#
    ElseIf IsEmpty(value) Or IsNull(value) Or value = vbNullString Then
        GetNumericValue = 0#
    Else
        GetNumericValue = CDbl(value)
    End If
End Function

Private Function ToDisplayValue(ByVal value As Variant) As Variant
    If IsEmpty(value) Or IsNull(value) Or value = vbNullString Then
        ToDisplayValue = vbNullString
    Else
        ToDisplayValue = value
    End If
End Function

Private Function ToSourceRowValue(ByVal value As Variant) As Variant
    If IsNumeric(value) Then
        If CLng(value) > 0 Then
            ToSourceRowValue = CLng(value)
        Else
            ToSourceRowValue = vbNullString
        End If
    Else
        ToSourceRowValue = vbNullString
    End If
End Function

Private Function FormatSourceAddress(ByVal filePath As Variant, ByVal rowNumber As Variant) As Variant
    If (IsEmpty(filePath) Or IsNull(filePath) Or filePath = vbNullString) Then
        FormatSourceAddress = vbNullString
    ElseIf Not IsNumeric(rowNumber) Then
        FormatSourceAddress = CStr(filePath)
    ElseIf CLng(rowNumber) <= 0 Then
        FormatSourceAddress = CStr(filePath)
    Else
        FormatSourceAddress = CStr(filePath) & " [Row " & CStr(CLng(rowNumber)) & "]"
    End If
End Function

Private Function GetField(ByVal fields As Variant, ByVal index As Long) As String
    If index <= UBound(fields) Then
        GetField = fields(index)
    Else
        GetField = vbNullString
    End If
End Function

Private Sub WriteHeaders(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim arr() As Variant
    Dim i As Long
    Dim count As Long
    count = UBound(headers) - LBound(headers) + 1
    ReDim arr(1 To 1, 1 To count)
    For i = LBound(headers) To UBound(headers)
        arr(1, i - LBound(headers) + 1) = headers(i)
    Next i
    ws.Range("A1").Resize(1, count).Value = arr
End Sub

Private Function EnsureOutputSheet(ByVal sheetName As String, ByVal headers As Variant) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    WriteHeaders ws, headers
    Set EnsureOutputSheet = ws
End Function

Private Function EnsureLogSheet(ByVal sheetName As String, ByVal headers As Variant) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    WriteHeaders ws, headers
    ws.Visible = xlSheetVisible
    Set EnsureLogSheet = ws
End Function

Private Sub HideWorksheet(ByVal ws As Worksheet)
    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub LogError(ByVal ws As Worksheet, ByVal filePath As String, ByVal lineNo As Long, _
                     ByVal rawTimestamp As String, ByVal message As String)
    Dim nextRow As Long
    Dim data() As Variant
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ReDim data(1 To 1, 1 To 4)
    data(1, 1) = filePath
    data(1, 2) = lineNo
    data(1, 3) = rawTimestamp
    data(1, 4) = message
    ws.Cells(nextRow, 1).Resize(1, 4).Value = data
End Sub

Private Sub LogSkipped(ByVal ws As Worksheet, ByVal filePath As String, ByVal reason As String)
    Dim nextRow As Long
    Dim data() As Variant
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ReDim data(1 To 1, 1 To 2)
    data(1, 1) = filePath
    data(1, 2) = reason
    ws.Cells(nextRow, 1).Resize(1, 2).Value = data
End Sub

Private Sub LogDuplicate(ByVal ws As Worksheet, ByVal filePath As String, ByVal timestampValue As Double, _
                         ByVal meterLabel As String, ByVal action As String)
    Dim nextRow As Long
    Dim data() As Variant
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ReDim data(1 To 1, 1 To 4)
    data(1, 1) = filePath
    data(1, 2) = Format$(timestampValue, "dd.mmm.yyyy hh:mm:ss")
    data(1, 3) = meterLabel
    data(1, 4) = action
    ws.Cells(nextRow, 1).Resize(1, 4).Value = data
End Sub

Private Sub ShowSummaryMessage()
    Dim message As String
    message = "Flash Record Summary:" & vbCrLf & _
              "Files processed: " & gStats.FilesProcessed & vbCrLf & _
              "Rows imported: " & gStats.RowsImported & vbCrLf & _
              "Duplicates: " & gStats.Duplicates & vbCrLf & _
              "Errors: " & gStats.Errors & vbCrLf & _
              "Skipped files: " & gStats.SkippedFiles
    MsgBox message, vbInformation
End Sub

Private Sub ResetStats(ByRef stats As SummaryStats)
    stats.FilesProcessed = 0
    stats.RowsImported = 0
    stats.Duplicates = 0
    stats.Errors = 0
    stats.SkippedFiles = 0
End Sub
