Attribute VB_Name = "DashboardUI"
Option Explicit

' ============================================================================
' Bank of England Loan Tape - Dashboard UI Module
' ============================================================================
' This module handles all user interface interactions from the Dashboard
' ============================================================================

' Main validation button procedure
Public Sub btnRunValidation_Click()
    ValidationEngine.RunFullValidation
End Sub

' View Error Report button
Public Sub btnViewErrorReport_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    
    ws.Activate
    ws.Range("A4").Select
    
    ' Apply filters if not already applied
    If Not ws.AutoFilterMode Then
        ws.Range("A4:I4").AutoFilter
    End If
End Sub

' Export errors to CSV
Public Sub btnExportErrors_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim fileDialog As FileDialog
    Dim csvContent As String
    Dim i As Long, j As Long
    Dim fNum As Integer
    
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    
    ' Check if there are errors to export
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 5 Then
        MsgBox "No errors to export. Run validation first.", vbInformation
        Exit Sub
    End If
    
    ' Get save location
    Set fileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With fileDialog
        .Title = "Export Error Report"
        .FilterIndex = 1
        .InitialFileName = "BoE_Validation_Errors_" & Format(Now, "YYYYMMDD_HHMMSS") & ".csv"
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub ' User cancelled
        End If
    End With
    
    ' Export to CSV
    On Error GoTo ErrorHandler
    
    fNum = FreeFile
    Open filePath For Output As #fNum
    
    ' Write headers
    For j = 1 To 9
        If j > 1 Then csvContent = csvContent & ","
        csvContent = csvContent & """" & ws.Cells(4, j).Value & """"
    Next j
    Print #fNum, csvContent
    
    ' Write data
    For i = 5 To lastRow
        csvContent = ""
        For j = 1 To 9
            If j > 1 Then csvContent = csvContent & ","
            csvContent = csvContent & """" & Replace(ws.Cells(i, j).Value, """", """""") & """"
        Next j
        Print #fNum, csvContent
    Next i
    
    Close #fNum
    
    MsgBox "Error report exported successfully to:" & vbCrLf & filePath, vbInformation
    Exit Sub
    
ErrorHandler:
    Close #fNum
    MsgBox "Error exporting file: " & Err.Description, vbCritical
End Sub

' Clear all data button
Public Sub btnClearAllData_Click()
    Dim response As VbMsgBoxResult
    Dim ws As Worksheet
    
    response = MsgBox("⚠ WARNING" & vbCrLf & vbCrLf & _
                     "This will clear ALL loan data and validation results." & vbCrLf & _
                     "This action cannot be undone." & vbCrLf & vbCrLf & _
                     "Do you want to proceed?", _
                     vbExclamation + vbYesNo, "Confirm Clear Data")
    
    If response = vbYes Then
        Application.ScreenUpdating = False
        
        ' Clear loan data
        Set ws = ThisWorkbook.Sheets("LoanData")
        If ws.Cells(5, 1).Value <> "" Then
            ws.Range("A5:HZ" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).ClearContents
        End If
        
        ' Clear error report
        Set ws = ThisWorkbook.Sheets("ErrorReport")
        If ws.Cells(5, 1).Value <> "" Then
            ws.Range("A5:I" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).ClearContents
        End If
        
        ' Reset dashboard
        ResetDashboard
        
        Application.ScreenUpdating = True
        
        MsgBox "All data cleared successfully.", vbInformation
        ThisWorkbook.Sheets("Dashboard").Activate
    End If
End Sub

' Reset dashboard to initial state
Private Sub ResetDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Reset summary values
    ws.Range("B9").Value = 0   ' Total Loans
    ws.Range("B10").Value = 0  ' Loans Validated
    ws.Range("B11").Value = 0  ' CRITICAL
    ws.Range("B12").Value = 0  ' WARNING
    ws.Range("B13").Value = 0  ' INFO
    ws.Range("B14").Value = 0  ' Total Issues
    
    ' Reset metrics
    ws.Range("F11").Value = "0%"  ' Data Completeness
    ws.Range("F12").Value = "Not Run"
    ws.Range("F12").Interior.ColorIndex = xlNone
    ws.Range("F12").Font.Color = RGB(0, 0, 0)
    ws.Range("F12").Font.Bold = False
    ws.Range("F13").Value = "Never"
    
    ' Reset category breakdown
    Dim i As Long
    For i = 18 To 26 ' Category rows
        ws.Cells(i, 2).Value = 0
        ws.Cells(i, 3).Value = 0
        ws.Cells(i, 4).Value = 0
        ws.Cells(i, 5).Value = 0
    Next i
End Sub

' Quick navigation procedures
Public Sub NavigateToLoanData()
    ThisWorkbook.Sheets("LoanData").Activate
    ThisWorkbook.Sheets("LoanData").Range("A5").Select
End Sub

Public Sub NavigateToErrorReport()
    ThisWorkbook.Sheets("ErrorReport").Activate
    ThisWorkbook.Sheets("ErrorReport").Range("A4").Select
End Sub

Public Sub NavigateToFieldDictionary()
    ThisWorkbook.Sheets("FieldDictionary").Activate
    ThisWorkbook.Sheets("FieldDictionary").Range("A1").Select
End Sub

' Data quality metrics calculation
Public Sub CalculateDataQuality()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalFields As Long, populatedFields As Long
    Dim i As Long, j As Long
    Dim completeness As Double
    
    Set ws = ThisWorkbook.Sheets("LoanData")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 5 Then Exit Sub
    
    totalFields = 0
    populatedFields = 0
    
    ' Count mandatory fields only
    For i = 5 To lastRow
        For j = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If ws.Cells(4, j).Value = "Mandatory" Then
                totalFields = totalFields + 1
                If Not IsEmpty(ws.Cells(i, j).Value) And ws.Cells(i, j).Value <> "" Then
                    populatedFields = populatedFields + 1
                End If
            End If
        Next j
    Next i
    
    If totalFields > 0 Then
        completeness = populatedFields / totalFields
        ThisWorkbook.Sheets("Dashboard").Range("F11").Value = Format(completeness, "0.0%")
    End If
End Sub

' Highlight errors in loan data sheet
Public Sub HighlightErrorsInData()
    Dim wsData As Worksheet, wsError As Worksheet
    Dim lastErrorRow As Long
    Dim i As Long
    Dim loanRow As Long
    Dim fieldsInvolved As String
    Dim fields() As String
    Dim field As Variant
    Dim col As Long
    
    Set wsData = ThisWorkbook.Sheets("LoanData")
    Set wsError = ThisWorkbook.Sheets("ErrorReport")
    
    ' Clear previous highlighting
    wsData.Cells.Interior.ColorIndex = xlNone
    
    lastErrorRow = wsError.Cells(wsError.Rows.Count, 1).End(xlUp).Row
    If lastErrorRow < 5 Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Process each error
    For i = 5 To lastErrorRow
        loanRow = wsError.Cells(i, 1).Value
        fieldsInvolved = wsError.Cells(i, 7).Value
        
        ' Split fields by comma
        fields = Split(fieldsInvolved, ",")
        
        For Each field In fields
            field = Trim(field)
            col = FindColumnByAR(wsData, CStr(field))
            
            If col > 0 Then
                ' Color based on severity
                Select Case wsError.Cells(i, 5).Value
                    Case "CRITICAL"
                        wsData.Cells(loanRow, col).Interior.Color = RGB(255, 199, 206)
                    Case "WARNING"
                        wsData.Cells(loanRow, col).Interior.Color = RGB(255, 235, 156)
                    Case "INFO"
                        wsData.Cells(loanRow, col).Interior.Color = RGB(198, 239, 206)
                End Select
            End If
        Next field
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Error cells highlighted in LoanData sheet." & vbCrLf & _
           "Red = CRITICAL, Yellow = WARNING, Green = INFO", vbInformation
End Sub

' Find column by AR code helper function
Private Function FindColumnByAR(ws As Worksheet, arCode As String) As Long
    Dim col As Long
    
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If UCase(Trim(ws.Cells(1, col).Value)) = UCase(arCode) Then
            FindColumnByAR = col
            Exit Function
        End If
    Next col
    
    FindColumnByAR = 0
End Function

' Create summary report
Public Sub GenerateSummaryReport()
    Dim ws As Worksheet
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim categoryCount As Object
    Dim severityCount As Object
    Dim category As Variant, Severity As Variant
    Dim rowNum As Long
    
    ' Check if errors exist
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 5 Then
        MsgBox "No errors to summarize. Run validation first.", vbInformation
        Exit Sub
    End If
    
    ' Create or clear summary sheet
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Sheets("ValidationSummary")
    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        summaryWs.Name = "ValidationSummary"
    Else
        summaryWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Create summary
    With summaryWs
        .Range("A1").Value = "VALIDATION SUMMARY REPORT"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A3").Value = "Generated: " & Format(Now, "DD-MM-YYYY HH:MM:SS")
        
        ' Overall statistics
        .Range("A5").Value = "Overall Statistics"
        .Range("A5").Font.Bold = True
        
        .Range("A6").Value = "Total Errors:"
        .Range("B6").Value = lastRow - 4
        
        ' Count by severity
        Set severityCount = CreateObject("Scripting.Dictionary")
        For i = 5 To lastRow
            Severity = ws.Cells(i, 5).Value
            If severityCount.exists(Severity) Then
                severityCount(Severity) = severityCount(Severity) + 1
            Else
                severityCount.Add Severity, 1
            End If
        Next i
        
        rowNum = 7
        For Each Severity In severityCount.keys
            .Cells(rowNum, 1).Value = Severity & " Errors:"
            .Cells(rowNum, 2).Value = severityCount(Severity)
            rowNum = rowNum + 1
        Next Severity
        
        ' Count by category
        rowNum = rowNum + 2
        .Cells(rowNum, 1).Value = "Errors by Category"
        .Cells(rowNum, 1).Font.Bold = True
        rowNum = rowNum + 1
        
        Set categoryCount = CreateObject("Scripting.Dictionary")
        For i = 5 To lastRow
            category = ws.Cells(i, 4).Value
            If categoryCount.exists(category) Then
                categoryCount(category) = categoryCount(category) + 1
            Else
                categoryCount.Add category, 1
            End If
        Next i
        
        For Each category In categoryCount.keys
            .Cells(rowNum, 1).Value = category
            .Cells(rowNum, 2).Value = categoryCount(category)
            rowNum = rowNum + 1
        Next category
        
        .Columns("A:B").AutoFit
    End With
    
    summaryWs.Activate
    MsgBox "Summary report generated in ValidationSummary sheet.", vbInformation
End Sub

' Filter error report by severity
Public Sub FilterBySeverity(Severity As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    
    ws.Activate
    
    ' Apply autofilter if not already applied
    If Not ws.AutoFilterMode Then
        ws.Range("A4:I4").AutoFilter
    End If
    
    ' Clear any existing filters
    ws.AutoFilter.ShowAllData
    
    ' Apply severity filter
    ws.Range("A4:I4").AutoFilter Field:=5, Criteria1:=Severity
    
    ws.Range("A4").Select
End Sub

' Quick filter buttons
Public Sub ShowCriticalErrorsOnly()
    FilterBySeverity "CRITICAL"
End Sub

Public Sub ShowWarningsOnly()
    FilterBySeverity "WARNING"
End Sub

Public Sub ShowAllErrors()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If
    
    ws.Activate
    ws.Range("A4").Select
End Sub
