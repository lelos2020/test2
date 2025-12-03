Attribute VB_Name = "UtilityFunctions"
Option Explicit

' ============================================================================
' Bank of England Loan Tape - Utility Functions
' ============================================================================
' Common utility functions used across the application
' ============================================================================

' Format currency with proper thousand separators
Public Function FormatCurrency(value As Variant) As String
    If IsNumeric(value) Then
        FormatCurrency = Format(value, "#,##0.00")
    Else
        FormatCurrency = ""
    End If
End Function

' Format percentage with one decimal
Public Function FormatPercentage(value As Variant) As String
    If IsNumeric(value) Then
        FormatPercentage = Format(value, "0.0") & "%"
    Else
        FormatPercentage = ""
    End If
End Function

' Validate date format (DD-MM-YYYY)
Public Function IsValidDate(dateStr As String) As Boolean
    Dim dateParts() As String
    Dim day As Integer, month As Integer, year As Integer
    
    On Error GoTo ErrorHandler
    
    If dateStr Like "##-##-####" Then
        dateParts = Split(dateStr, "-")
        day = CInt(dateParts(0))
        month = CInt(dateParts(1))
        year = CInt(dateParts(2))
        
        If month >= 1 And month <= 12 And day >= 1 And day <= 31 Then
            IsValidDate = True
        Else
            IsValidDate = False
        End If
    Else
        IsValidDate = False
    End If
    Exit Function
    
ErrorHandler:
    IsValidDate = False
End Function

' Convert Excel date to DD-MM-YYYY format
Public Function ConvertDateToDDMMYYYY(dateValue As Variant) As String
    If IsDate(dateValue) Then
        ConvertDateToDDMMYYYY = Format(dateValue, "DD-MM-YYYY")
    Else
        ConvertDateToDDMMYYYY = ""
    End If
End Function

' Validate numeric value is within range
Public Function IsInRange(value As Variant, minVal As Double, maxVal As Double) As Boolean
    If IsNumeric(value) Then
        IsInRange = (CDbl(value) >= minVal And CDbl(value) <= maxVal)
    Else
        IsInRange = False
    End If
End Function

' Clean and trim string
Public Function CleanString(str As String) As String
    CleanString = Trim(Replace(Replace(str, vbCr, ""), vbLf, ""))
End Function

' Check if value is in a list (array)
Public Function ValueInList(value As Variant, list As Variant) As Boolean
    Dim item As Variant
    
    For Each item In list
        If CStr(value) = CStr(item) Then
            ValueInList = True
            Exit Function
        End If
    Next item
    
    ValueInList = False
End Function

' Get column letter from column number
Public Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

' Count non-empty cells in range
Public Function CountNonEmpty(rng As Range) As Long
    Dim cell As Range
    Dim count As Long
    
    count = 0
    For Each cell In rng
        If Not IsEmpty(cell.Value) And cell.Value <> "" Then
            count = count + 1
        End If
    Next cell
    
    CountNonEmpty = count
End Function

' Calculate percentage
Public Function CalculatePercentage(part As Variant, whole As Variant) As Double
    If IsNumeric(part) And IsNumeric(whole) And CDbl(whole) <> 0 Then
        CalculatePercentage = (CDbl(part) / CDbl(whole)) * 100
    Else
        CalculatePercentage = 0
    End If
End Function

' Safe division (returns 0 if divide by zero)
Public Function SafeDivide(numerator As Variant, denominator As Variant) As Double
    If IsNumeric(numerator) And IsNumeric(denominator) Then
        If CDbl(denominator) <> 0 Then
            SafeDivide = CDbl(numerator) / CDbl(denominator)
        Else
            SafeDivide = 0
        End If
    Else
        SafeDivide = 0
    End If
End Function

' Check if sheet exists
Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Get last row with data in column
Public Function GetLastRow(ws As Worksheet, col As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Get last column with data in row
Public Function GetLastColumn(ws As Worksheet, Row As Long) As Long
    GetLastColumn = ws.Cells(Row, ws.Columns.Count).End(xlToLeft).Column
End Function

' Clear all filters on a sheet
Public Sub ClearAllFilters(ws As Worksheet)
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If
End Sub

' Apply autofilter to range
Public Sub ApplyAutoFilter(ws As Worksheet, startRow As Long, endCol As Long)
    Dim rng As Range
    
    Set rng = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, endCol))
    
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If
    
    rng.AutoFilter
End Sub

' Color cell based on value
Public Sub ColorCodeCell(cell As Range, value As Variant, lowThreshold As Double, _
                        highThreshold As Double, Optional reverse As Boolean = False)
    
    If Not IsNumeric(value) Then Exit Sub
    
    Dim numValue As Double
    numValue = CDbl(value)
    
    If reverse Then
        ' Higher is worse
        If numValue >= highThreshold Then
            cell.Interior.Color = RGB(255, 199, 206) ' Red
        ElseIf numValue >= lowThreshold Then
            cell.Interior.Color = RGB(255, 235, 156) ' Yellow
        Else
            cell.Interior.Color = RGB(198, 239, 206) ' Green
        End If
    Else
        ' Higher is better
        If numValue >= highThreshold Then
            cell.Interior.Color = RGB(198, 239, 206) ' Green
        ElseIf numValue >= lowThreshold Then
            cell.Interior.Color = RGB(255, 235, 156) ' Yellow
        Else
            cell.Interior.Color = RGB(255, 199, 206) ' Red
        End If
    End If
End Sub

' Create backup of current data
Public Function CreateBackup() As Boolean
    Dim backupPath As String
    Dim timestamp As String
    
    On Error GoTo ErrorHandler
    
    timestamp = Format(Now, "YYYYMMDD_HHMMSS")
    backupPath = ThisWorkbook.Path & "\Backup_" & timestamp & ".xlsx"
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs backupPath
    Application.DisplayAlerts = True
    
    CreateBackup = True
    MsgBox "Backup created: " & backupPath, vbInformation
    Exit Function
    
ErrorHandler:
    CreateBackup = False
    MsgBox "Error creating backup: " & Err.Description, vbCritical
End Function

' Export sheet to CSV
Public Function ExportSheetToCSV(ws As Worksheet, Optional fileName As String = "") As Boolean
    Dim filePath As String
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim csvContent As String
    Dim fNum As Integer
    
    On Error GoTo ErrorHandler
    
    lastRow = GetLastRow(ws, 1)
    lastCol = GetLastColumn(ws, 1)
    
    If fileName = "" Then
        fileName = ws.Name & "_" & Format(Now, "YYYYMMDD_HHMMSS") & ".csv"
    End If
    
    filePath = ThisWorkbook.Path & "\" & fileName
    
    fNum = FreeFile
    Open filePath For Output As #fNum
    
    For i = 1 To lastRow
        csvContent = ""
        For j = 1 To lastCol
            If j > 1 Then csvContent = csvContent & ","
            csvContent = csvContent & """" & Replace(ws.Cells(i, j).Value, """", """""") & """"
        Next j
        Print #fNum, csvContent
    Next i
    
    Close #fNum
    
    ExportSheetToCSV = True
    Exit Function
    
ErrorHandler:
    Close #fNum
    ExportSheetToCSV = False
End Function

' Sanitize text for CSV export
Public Function SanitizeForCSV(text As String) As String
    Dim result As String
    result = text
    
    ' Replace quotes
    result = Replace(result, """", """""")
    
    ' Remove line breaks
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    
    ' Trim
    result = Trim(result)
    
    SanitizeForCSV = result
End Function

' Log message to immediate window (for debugging)
Public Sub LogDebug(msg As String)
    Debug.Print Format(Now, "HH:MM:SS") & " | " & msg
End Sub

' Show progress indicator
Public Sub ShowProgress(current As Long, total As Long, Optional taskName As String = "Processing")
    Dim percentage As Double
    
    If total > 0 Then
        percentage = (current / total) * 100
        Application.StatusBar = taskName & ": " & current & " of " & total & " (" & Format(percentage, "0") & "%)"
        DoEvents
    End If
End Sub

' Clear progress indicator
Public Sub ClearProgress()
    Application.StatusBar = False
End Sub

' Validate AR code format
Public Function IsValidARCode(arCode As String) As Boolean
    ' AR codes should be in format "AR" followed by 1-3 digits
    IsValidARCode = (arCode Like "AR#" Or arCode Like "AR##" Or arCode Like "AR###")
End Function

' Get field name for AR code from dictionary
Public Function GetFieldName(arCode As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("FieldDictionary")
    lastRow = GetLastRow(ws, 1)
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = arCode Then
            GetFieldName = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    GetFieldName = ""
End Function

' Get data type for AR code
Public Function GetDataType(arCode As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("FieldDictionary")
    lastRow = GetLastRow(ws, 1)
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = arCode Then
            GetDataType = ws.Cells(i, 5).Value
            Exit Function
        End If
    Next i
    
    GetDataType = ""
End Function

' Check if field is mandatory
Public Function IsMandatoryField(arCode As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("FieldDictionary")
    lastRow = GetLastRow(ws, 1)
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = arCode Then
            IsMandatoryField = (ws.Cells(i, 3).Value = "Mandatory")
            Exit Function
        End If
    Next i
    
    IsMandatoryField = False
End Function

' Round to specified decimal places
Public Function RoundTo(value As Variant, decimalPlaces As Integer) As Double
    If IsNumeric(value) Then
        RoundTo = Round(CDbl(value), decimalPlaces)
    Else
        RoundTo = 0
    End If
End Function

' Calculate LTV
Public Function CalculateLTV(balance As Variant, valuation As Variant) As Double
    If IsNumeric(balance) And IsNumeric(valuation) And CDbl(valuation) > 0 Then
        CalculateLTV = (CDbl(balance) / CDbl(valuation)) * 100
    Else
        CalculateLTV = 0
    End If
End Function

' Calculate DSCR (Debt Service Coverage Ratio)
Public Function CalculateDSCR(rentalIncome As Variant, payment As Variant, _
                              Optional stressRate As Double = 1.25) As Double
    If IsNumeric(rentalIncome) And IsNumeric(payment) And CDbl(payment) > 0 Then
        CalculateDSCR = ((CDbl(rentalIncome) * stressRate) / (CDbl(payment) * 12)) * 100
    Else
        CalculateDSCR = 0
    End If
End Function
