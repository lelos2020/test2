Attribute VB_Name = "ModuleStratification"
'============================================================================
' DYNAMIC PORTFOLIO STRATIFICATION TOOL
' Version: 1.0
' Purpose: Fully automated, dynamic loan tape stratification engine
' Author: Created for UK Securitisation Structuring
' Date: January 2025
'============================================================================

Option Explicit

'============================================================================
' MAIN STRATIFICATION ENGINE
'============================================================================

Sub GenerateAllStratifications()
    '
    ' Main procedure to generate all portfolio stratifications
    ' Called from Control Panel button
    '
    
    Dim wsControl As Worksheet
    Dim wsFields As Worksheet
    Dim wsBuckets As Worksheet
    Dim wsData As Worksheet
    Dim wsStrat As Worksheet
    Dim wsStats As Worksheet
    
    Dim startTime As Double
    Dim fieldCount As Long
    Dim lastRow As Long
    
    ' Start timer
    startTime = Timer
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Set worksheet references
    Set wsControl = ThisWorkbook.Sheets("Control Panel")
    Set wsFields = ThisWorkbook.Sheets("Field Definitions")
    Set wsBuckets = ThisWorkbook.Sheets("Bucket Definitions")
    Set wsData = ThisWorkbook.Sheets("Loan Tape Data")
    Set wsStrat = ThisWorkbook.Sheets("Stratifications")
    Set wsStats = ThisWorkbook.Sheets("Summary Stats")
    
    ' Validate data exists
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No loan data found. Please load data first.", vbExclamation, "No Data"
        GoTo CleanExit
    End If
    
    ' Clear previous stratifications
    wsStrat.Cells.Clear
    wsStats.Cells.Clear
    
    ' Display progress
    Application.StatusBar = "Generating stratifications... Please wait..."
    
    ' Generate summary statistics first
    Call GenerateSummaryStats
    
    ' Process each field for stratification
    Call ProcessAllFields
    
    ' Update control panel
    wsControl.Range("B12").Value = Now
    wsControl.Range("B2").Value = "DATA LOADED"
    wsControl.Range("B2").Font.Color = RGB(0, 176, 80)
    
    ' Success message
    MsgBox "Stratifications generated successfully!" & vbCrLf & _
           "Processing time: " & Format(Timer - startTime, "0.00") & " seconds", _
           vbInformation, "Success"
    
CleanExit:
    ' Restore settings
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Error generating stratifications: " & Err.Description, vbCritical, "Error"
End Sub

'============================================================================
' DATA LOADING PROCEDURES
'============================================================================

Sub LoadLoanTapeData()
    '
    ' Load loan tape data from external file (CSV or Excel)
    '
    
    Dim fd As FileDialog
    Dim fileName As String
    Dim wsData As Worksheet
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set wsData = ThisWorkbook.Sheets("Loan Tape Data")
    
    ' Open file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Loan Tape Data File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls"
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            fileName = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Loading data..."
    
    ' Clear existing data
    wsData.Cells.Clear
    
    ' Load data based on file type
    If LCase(Right(fileName, 4)) = ".csv" Then
        ' Load CSV file
        Call LoadCSVFile(fileName, wsData)
    Else
        ' Load Excel file
        Set wbSource = Workbooks.Open(fileName, ReadOnly:=True)
        Set wsSource = wbSource.Sheets(1)
        
        ' Find data range
        lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
        lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
        
        ' Copy data
        wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol)).Copy
        wsData.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        
        wbSource.Close False
    End If
    
    ' Auto-detect fields and update Field Definitions
    Call AutoDetectFields
    
    ' Format data sheet
    Call FormatDataSheet
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Data loaded successfully!" & vbCrLf & _
           "Rows: " & (wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row - 1), _
           vbInformation, "Data Loaded"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error loading data: " & Err.Description, vbCritical, "Error"
End Sub

Sub LoadCSVFile(fileName As String, wsTarget As Worksheet)
    '
    ' Load CSV file into target worksheet
    '
    
    Dim fso As Object
    Dim ts As Object
    Dim lineText As String
    Dim dataArray() As String
    Dim row As Long, col As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(fileName, 1) ' 1 = ForReading
    
    row = 1
    Do While Not ts.AtEndOfStream
        lineText = ts.ReadLine
        dataArray = Split(lineText, ",")
        
        For col = 0 To UBound(dataArray)
            wsTarget.Cells(row, col + 1).Value = dataArray(col)
        Next col
        
        row = row + 1
    Loop
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

Sub AutoDetectFields()
    '
    ' Automatically detect fields from loaded data and populate Field Definitions
    '
    
    Dim wsData As Worksheet
    Dim wsFields As Worksheet
    Dim lastCol As Long
    Dim i As Long, row As Long
    Dim fieldName As String
    Dim dataType As String
    Dim sampleValue As Variant
    
    Set wsData = ThisWorkbook.Sheets("Loan Tape Data")
    Set wsFields = ThisWorkbook.Sheets("Field Definitions")
    
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    ' Clear existing field definitions (except header)
    If wsFields.Cells(2, 1).Value <> "" Then
        wsFields.Rows("2:" & wsFields.Rows.Count).ClearContents
    End If
    
    row = 2
    For i = 1 To lastCol
        fieldName = wsData.Cells(1, i).Value
        
        ' Detect data type from first non-empty value
        sampleValue = wsData.Cells(2, i).Value
        If IsNumeric(sampleValue) Then
            dataType = "Numeric"
        ElseIf IsDate(sampleValue) Then
            dataType = "Date"
        Else
            dataType = "Text"
        End If
        
        ' Populate field definition
        wsFields.Cells(row, 1).Value = row - 1 ' Index
        wsFields.Cells(row, 2).Value = fieldName
        wsFields.Cells(row, 3).Value = dataType
        wsFields.Cells(row, 4).Value = "Auto-detected from data"
        wsFields.Cells(row, 5).Value = "YES"
        
        ' Set default weighted average field for numeric fields
        If dataType = "Numeric" And InStr(1, LCase(fieldName), "balance") > 0 Then
            ' This is likely the balance field - don't set itself as weight
            wsFields.Cells(row, 6).Value = ""
        ElseIf dataType = "Numeric" Then
            ' For other numeric fields, use first balance field as weight
            wsFields.Cells(row, 6).Value = FindBalanceField(wsData)
        End If
        
        row = row + 1
    Next i
End Sub

Function FindBalanceField(wsData As Worksheet) As String
    '
    ' Find the balance field in the data (for weighted averages)
    '
    
    Dim lastCol As Long
    Dim i As Long
    Dim fieldName As String
    
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        fieldName = LCase(wsData.Cells(1, i).Value)
        If InStr(1, fieldName, "current balance") > 0 Or _
           InStr(1, fieldName, "outstanding") > 0 Or _
           InStr(1, fieldName, "balance") > 0 Then
            FindBalanceField = wsData.Cells(1, i).Value
            Exit Function
        End If
    Next i
    
    ' Default to first numeric column
    FindBalanceField = wsData.Cells(1, 1).Value
End Function

'============================================================================
' STRATIFICATION PROCESSING
'============================================================================

Sub ProcessAllFields()
    '
    ' Process all active fields and generate stratifications
    '
    
    Dim wsFields As Worksheet
    Dim wsBuckets As Worksheet
    Dim wsStrat As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fieldName As String
    Dim includeField As String
    Dim outputRow As Long
    
    Set wsFields = ThisWorkbook.Sheets("Field Definitions")
    Set wsBuckets = ThisWorkbook.Sheets("Bucket Definitions")
    Set wsStrat = ThisWorkbook.Sheets("Stratifications")
    
    lastRow = wsFields.Cells(wsFields.Rows.Count, 2).End(xlUp).Row
    outputRow = 1
    
    ' Loop through each field
    For i = 2 To lastRow
        fieldName = wsFields.Cells(i, 2).Value
        includeField = UCase(wsFields.Cells(i, 5).Value)
        
        If includeField = "YES" And fieldName <> "" Then
            Application.StatusBar = "Processing: " & fieldName
            
            ' Generate stratification for this field
            outputRow = GenerateFieldStratification(fieldName, outputRow)
            outputRow = outputRow + 3 ' Add spacing between stratifications
        End If
    Next i
    
    ' Format stratifications sheet
    Call FormatStratificationsSheet
End Sub

Function GenerateFieldStratification(fieldName As String, startRow As Long) As Long
    '
    ' Generate stratification for a specific field
    ' Returns: Next available row for output
    '
    
    Dim wsData As Worksheet
    Dim wsBuckets As Worksheet
    Dim wsStrat As Worksheet
    Dim wsFields As Worksheet
    
    Dim bucketArray() As Variant
    Dim dataCol As Long
    Dim bucketCount As Long
    Dim dataType As String
    Dim weightField As String
    Dim weightCol As Long
    
    Set wsData = ThisWorkbook.Sheets("Loan Tape Data")
    Set wsBuckets = ThisWorkbook.Sheets("Bucket Definitions")
    Set wsStrat = ThisWorkbook.Sheets("Stratifications")
    Set wsFields = ThisWorkbook.Sheets("Field Definitions")
    
    ' Find data column
    dataCol = FindColumnByName(wsData, fieldName)
    If dataCol = 0 Then
        GenerateFieldStratification = startRow
        Exit Function
    End If
    
    ' Get field data type and weight field
    dataType = GetFieldDataType(fieldName)
    weightField = GetWeightField(fieldName)
    If weightField <> "" Then
        weightCol = FindColumnByName(wsData, weightField)
    End If
    
    ' Get buckets for this field
    bucketArray = GetBucketsForField(fieldName)
    
    If UBound(bucketArray, 1) = 0 Then
        ' No buckets defined - create default buckets
        If dataType = "Text" Then
            bucketArray = CreateTextBuckets(wsData, dataCol)
        ElseIf dataType = "Numeric" Then
            bucketArray = CreateNumericBuckets(wsData, dataCol)
        End If
    End If
    
    ' Write stratification header
    Call WriteStratificationHeader(wsStrat, startRow, fieldName)
    
    ' Calculate and write stratification data
    Call CalculateStratification(wsData, wsStrat, dataCol, weightCol, bucketArray, startRow + 2)
    
    GenerateFieldStratification = startRow + UBound(bucketArray, 1) + 4
End Function

Sub WriteStratificationHeader(ws As Worksheet, row As Long, fieldName As String)
    '
    ' Write header for a stratification table
    '
    
    ' Title
    ws.Cells(row, 1).Value = "STRATIFICATION: " & UCase(fieldName)
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Size = 12
    ws.Cells(row, 1).Font.Color = RGB(54, 96, 146)
    
    ' Column headers
    ws.Cells(row + 1, 1).Value = "Bucket"
    ws.Cells(row + 1, 2).Value = "# of Loans"
    ws.Cells(row + 1, 3).Value = "Balance (£)"
    ws.Cells(row + 1, 4).Value = "% of Count"
    ws.Cells(row + 1, 5).Value = "% of Balance"
    ws.Cells(row + 1, 6).Value = "Avg Balance"
    ws.Cells(row + 1, 7).Value = "Wtd Avg Rate"
    ws.Cells(row + 1, 8).Value = "Wtd Avg LTV"
    ws.Cells(row + 1, 9).Value = "Wtd Avg Term"
    
    ' Format headers
    Dim i As Long
    For i = 1 To 9
        With ws.Cells(row + 1, i)
            .Font.Bold = True
            .Interior.Color = RGB(54, 96, 146)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
    Next i
End Sub

Sub CalculateStratification(wsData As Worksheet, wsStrat As Worksheet, _
                            dataCol As Long, weightCol As Long, _
                            bucketArray As Variant, startRow As Long)
    '
    ' Calculate stratification metrics for each bucket
    '
    
    Dim lastRow As Long
    Dim totalCount As Long, totalBalance As Double
    Dim i As Long, j As Long
    Dim bucketLabel As String
    Dim minVal As Double, maxVal As Double
    Dim cellValue As Variant
    Dim bucketCount As Long
    Dim bucketBalance As Double
    Dim outputRow As Long
    Dim rateCol As Long, ltvCol As Long, termCol As Long
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Find additional columns for weighted averages
    rateCol = FindColumnByName(wsData, "Interest Rate")
    ltvCol = FindColumnByName(wsData, "LTV")
    termCol = FindColumnByName(wsData, "Remaining Term")
    
    ' Calculate totals
    totalCount = lastRow - 1
    If weightCol > 0 Then
        totalBalance = Application.WorksheetFunction.Sum(wsData.Range(wsData.Cells(2, weightCol), wsData.Cells(lastRow, weightCol)))
    End If
    
    outputRow = startRow
    
    ' Process each bucket
    For i = LBound(bucketArray, 1) To UBound(bucketArray, 1)
        bucketLabel = bucketArray(i, 0)
        
        ' Initialize counters
        bucketCount = 0
        bucketBalance = 0
        
        ' Determine bucket type and calculate
        If IsNumeric(bucketArray(i, 1)) Then
            ' Numeric range bucket
            minVal = CDbl(bucketArray(i, 1))
            maxVal = CDbl(bucketArray(i, 2))
            
            ' Count loans in bucket
            For j = 2 To lastRow
                cellValue = wsData.Cells(j, dataCol).Value
                If IsNumeric(cellValue) Then
                    If cellValue >= minVal And cellValue < maxVal Then
                        bucketCount = bucketCount + 1
                        If weightCol > 0 Then
                            bucketBalance = bucketBalance + wsData.Cells(j, weightCol).Value
                        End If
                    End If
                End If
            Next j
        Else
            ' Text bucket
            For j = 2 To lastRow
                cellValue = wsData.Cells(j, dataCol).Value
                If CStr(cellValue) = bucketLabel Then
                    bucketCount = bucketCount + 1
                    If weightCol > 0 Then
                        bucketBalance = bucketBalance + wsData.Cells(j, weightCol).Value
                    End If
                End If
            Next j
        End If
        
        ' Write results
        wsStrat.Cells(outputRow, 1).Value = bucketLabel
        wsStrat.Cells(outputRow, 2).Value = bucketCount
        wsStrat.Cells(outputRow, 3).Value = bucketBalance
        
        If totalCount > 0 Then
            wsStrat.Cells(outputRow, 4).Value = bucketCount / totalCount
            wsStrat.Cells(outputRow, 4).NumberFormat = "0.00%"
        End If
        
        If totalBalance > 0 Then
            wsStrat.Cells(outputRow, 5).Value = bucketBalance / totalBalance
            wsStrat.Cells(outputRow, 5).NumberFormat = "0.00%"
        End If
        
        If bucketCount > 0 Then
            wsStrat.Cells(outputRow, 6).Value = bucketBalance / bucketCount
        End If
        
        ' Format numbers
        wsStrat.Cells(outputRow, 2).NumberFormat = "#,##0"
        wsStrat.Cells(outputRow, 3).NumberFormat = "£#,##0"
        wsStrat.Cells(outputRow, 6).NumberFormat = "£#,##0"
        
        outputRow = outputRow + 1
    Next i
    
    ' Add total row
    wsStrat.Cells(outputRow, 1).Value = "TOTAL"
    wsStrat.Cells(outputRow, 1).Font.Bold = True
    wsStrat.Cells(outputRow, 2).Value = totalCount
    wsStrat.Cells(outputRow, 3).Value = totalBalance
    wsStrat.Cells(outputRow, 4).Value = 1
    wsStrat.Cells(outputRow, 5).Value = 1
    
    wsStrat.Cells(outputRow, 2).NumberFormat = "#,##0"
    wsStrat.Cells(outputRow, 3).NumberFormat = "£#,##0"
    wsStrat.Cells(outputRow, 4).NumberFormat = "0.00%"
    wsStrat.Cells(outputRow, 5).NumberFormat = "0.00%"
End Sub

'============================================================================
' BUCKET CREATION FUNCTIONS
'============================================================================

Function GetBucketsForField(fieldName As String) As Variant
    '
    ' Get bucket definitions for a specific field
    '
    
    Dim wsBuckets As Worksheet
    Dim lastRow As Long
    Dim i As Long, bucketIdx As Long
    Dim buckets() As Variant
    
    Set wsBuckets = ThisWorkbook.Sheets("Bucket Definitions")
    lastRow = wsBuckets.Cells(wsBuckets.Rows.Count, 2).End(xlUp).Row
    
    ' Count matching buckets
    bucketIdx = 0
    For i = 2 To lastRow
        If wsBuckets.Cells(i, 2).Value = fieldName Then
            bucketIdx = bucketIdx + 1
        End If
    Next i
    
    If bucketIdx = 0 Then
        ' Return empty array
        ReDim buckets(0, 0)
        GetBucketsForField = buckets
        Exit Function
    End If
    
    ' Build bucket array
    ReDim buckets(1 To bucketIdx, 0 To 2)
    bucketIdx = 0
    
    For i = 2 To lastRow
        If wsBuckets.Cells(i, 2).Value = fieldName Then
            bucketIdx = bucketIdx + 1
            buckets(bucketIdx, 0) = wsBuckets.Cells(i, 6).Value ' Label
            buckets(bucketIdx, 1) = wsBuckets.Cells(i, 4).Value ' Min
            buckets(bucketIdx, 2) = wsBuckets.Cells(i, 5).Value ' Max
        End If
    Next i
    
    GetBucketsForField = buckets
End Function

Function CreateTextBuckets(wsData As Worksheet, dataCol As Long) As Variant
    '
    ' Auto-create buckets for text field based on unique values
    '
    
    Dim lastRow As Long
    Dim uniqueDict As Object
    Dim i As Long
    Dim cellValue As String
    Dim buckets() As Variant
    Dim bucketIdx As Long
    
    Set uniqueDict = CreateObject("Scripting.Dictionary")
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Find unique values
    For i = 2 To lastRow
        cellValue = CStr(wsData.Cells(i, dataCol).Value)
        If cellValue <> "" And Not uniqueDict.Exists(cellValue) Then
            uniqueDict.Add cellValue, 1
        End If
    Next i
    
    ' Create bucket array
    ReDim buckets(1 To uniqueDict.Count, 0 To 2)
    bucketIdx = 0
    
    Dim key As Variant
    For Each key In uniqueDict.Keys
        bucketIdx = bucketIdx + 1
        buckets(bucketIdx, 0) = key
        buckets(bucketIdx, 1) = ""
        buckets(bucketIdx, 2) = ""
    Next key
    
    CreateTextBuckets = buckets
    Set uniqueDict = Nothing
End Function

Function CreateNumericBuckets(wsData As Worksheet, dataCol As Long) As Variant
    '
    ' Auto-create buckets for numeric field based on data distribution
    '
    
    Dim lastRow As Long
    Dim minVal As Double, maxVal As Double
    Dim rangeSize As Double
    Dim buckets() As Variant
    Dim i As Long
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Find min and max
    minVal = Application.WorksheetFunction.Min(wsData.Range(wsData.Cells(2, dataCol), wsData.Cells(lastRow, dataCol)))
    maxVal = Application.WorksheetFunction.Max(wsData.Range(wsData.Cells(2, dataCol), wsData.Cells(lastRow, dataCol)))
    
    ' Create 5 buckets
    rangeSize = (maxVal - minVal) / 5
    ReDim buckets(1 To 5, 0 To 2)
    
    For i = 1 To 5
        buckets(i, 0) = Format(minVal + (i - 1) * rangeSize, "#,##0") & " - " & Format(minVal + i * rangeSize, "#,##0")
        buckets(i, 1) = minVal + (i - 1) * rangeSize
        buckets(i, 2) = minVal + i * rangeSize
    Next i
    
    CreateNumericBuckets = buckets
End Function

'============================================================================
' SUMMARY STATISTICS
'============================================================================

Sub GenerateSummaryStats()
    '
    ' Generate portfolio summary statistics
    '
    
    Dim wsData As Worksheet
    Dim wsStats As Worksheet
    Dim lastRow As Long
    Dim balanceCol As Long, rateCol As Long, ltvCol As Long, termCol As Long
    Dim row As Long
    
    Set wsData = ThisWorkbook.Sheets("Loan Tape Data")
    Set wsStats = ThisWorkbook.Sheets("Summary Stats")
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' Find key columns
    balanceCol = FindColumnByName(wsData, "Current Balance")
    rateCol = FindColumnByName(wsData, "Interest Rate")
    ltvCol = FindColumnByName(wsData, "LTV")
    termCol = FindColumnByName(wsData, "Remaining Term")
    
    ' Clear previous stats
    wsStats.Cells.Clear
    
    ' Title
    wsStats.Range("A1").Value = "PORTFOLIO SUMMARY STATISTICS"
    wsStats.Range("A1").Font.Bold = True
    wsStats.Range("A1").Font.Size = 14
    wsStats.Range("A1").Font.Color = RGB(54, 96, 146)
    
    row = 3
    
    ' Number of loans
    wsStats.Cells(row, 1).Value = "Number of Loans:"
    wsStats.Cells(row, 2).Value = lastRow - 1
    wsStats.Cells(row, 2).NumberFormat = "#,##0"
    row = row + 1
    
    ' Total balance
    If balanceCol > 0 Then
        wsStats.Cells(row, 1).Value = "Total Balance:"
        wsStats.Cells(row, 2).Formula = "=SUM('Loan Tape Data'!" & _
            Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & ")"
        wsStats.Cells(row, 2).NumberFormat = "£#,##0"
        row = row + 1
        
        ' Average balance
        wsStats.Cells(row, 1).Value = "Average Balance:"
        wsStats.Cells(row, 2).Formula = "=AVERAGE('Loan Tape Data'!" & _
            Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & ")"
        wsStats.Cells(row, 2).NumberFormat = "£#,##0"
        row = row + 1
    End If
    
    ' Weighted average rate
    If rateCol > 0 And balanceCol > 0 Then
        wsStats.Cells(row, 1).Value = "Weighted Avg Rate:"
        wsStats.Cells(row, 2).Formula = "=SUMPRODUCT('Loan Tape Data'!" & _
            Range(Cells(2, rateCol), Cells(lastRow, rateCol)).Address(External:=False) & "," & _
            "'Loan Tape Data'!" & Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & _
            ")/SUM('Loan Tape Data'!" & Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & ")"
        wsStats.Cells(row, 2).NumberFormat = "0.00%"
        row = row + 1
    End If
    
    ' Weighted average LTV
    If ltvCol > 0 And balanceCol > 0 Then
        wsStats.Cells(row, 1).Value = "Weighted Avg LTV:"
        wsStats.Cells(row, 2).Formula = "=SUMPRODUCT('Loan Tape Data'!" & _
            Range(Cells(2, ltvCol), Cells(lastRow, ltvCol)).Address(External:=False) & "," & _
            "'Loan Tape Data'!" & Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & _
            ")/SUM('Loan Tape Data'!" & Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & ")"
        wsStats.Cells(row, 2).NumberFormat = "0.00%"
        row = row + 1
    End If
    
    ' Weighted average term
    If termCol > 0 And balanceCol > 0 Then
        wsStats.Cells(row, 1).Value = "Weighted Avg Term (months):"
        wsStats.Cells(row, 2).Formula = "=SUMPRODUCT('Loan Tape Data'!" & _
            Range(Cells(2, termCol), Cells(lastRow, termCol)).Address(External:=False) & "," & _
            "'Loan Tape Data'!" & Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & _
            ")/SUM('Loan Tape Data'!" & Range(Cells(2, balanceCol), Cells(lastRow, balanceCol)).Address(External:=False) & ")"
        wsStats.Cells(row, 2).NumberFormat = "0.0"
        row = row + 1
    End If
    
    ' Format labels
    wsStats.Range("A3:A" & row - 1).Font.Bold = True
    wsStats.Columns("A:A").AutoFit
    wsStats.Columns("B:B").AutoFit
End Sub

'============================================================================
' HELPER FUNCTIONS
'============================================================================

Function FindColumnByName(ws As Worksheet, columnName As String) As Long
    '
    ' Find column number by header name
    '
    
    Dim lastCol As Long
    Dim i As Long
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If LCase(Trim(ws.Cells(1, i).Value)) = LCase(Trim(columnName)) Then
            FindColumnByName = i
            Exit Function
        End If
    Next i
    
    FindColumnByName = 0
End Function

Function GetFieldDataType(fieldName As String) As String
    '
    ' Get data type for a field from Field Definitions
    '
    
    Dim wsFields As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wsFields = ThisWorkbook.Sheets("Field Definitions")
    lastRow = wsFields.Cells(wsFields.Rows.Count, 2).End(xlUp).Row
    
    For i = 2 To lastRow
        If wsFields.Cells(i, 2).Value = fieldName Then
            GetFieldDataType = wsFields.Cells(i, 3).Value
            Exit Function
        End If
    Next i
    
    GetFieldDataType = "Text"
End Function

Function GetWeightField(fieldName As String) As String
    '
    ' Get weighted average field for a field from Field Definitions
    '
    
    Dim wsFields As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wsFields = ThisWorkbook.Sheets("Field Definitions")
    lastRow = wsFields.Cells(wsFields.Rows.Count, 2).End(xlUp).Row
    
    For i = 2 To lastRow
        If wsFields.Cells(i, 2).Value = fieldName Then
            GetWeightField = wsFields.Cells(i, 6).Value
            Exit Function
        End If
    Next i
    
    GetWeightField = ""
End Function

Sub FormatDataSheet()
    '
    ' Format the Loan Tape Data sheet
    '
    
    Dim wsData As Worksheet
    Dim lastCol As Long
    
    Set wsData = ThisWorkbook.Sheets("Loan Tape Data")
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    
    ' Format header row
    With wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(54, 96, 146)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Auto-fit columns
    wsData.Columns.AutoFit
End Sub

Sub FormatStratificationsSheet()
    '
    ' Format the Stratifications sheet
    '
    
    Dim wsStrat As Worksheet
    Set wsStrat = ThisWorkbook.Sheets("Stratifications")
    
    ' Auto-fit columns
    wsStrat.Columns.AutoFit
    
    ' Freeze top row if needed
    wsStrat.Activate
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 0
    ActiveWindow.FreezePanes = False
End Sub

'============================================================================
' EXPORT FUNCTIONS
'============================================================================

Sub ExportToPDF()
    '
    ' Export stratifications to PDF
    '
    
    Dim wsStrat As Worksheet
    Dim fd As FileDialog
    Dim fileName As String
    
    Set wsStrat = ThisWorkbook.Sheets("Stratifications")
    
    ' Check if data exists
    If wsStrat.Range("A1").Value = "" Then
        MsgBox "No stratifications to export. Please generate stratifications first.", vbExclamation
        Exit Sub
    End If
    
    ' File dialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Save Stratifications as PDF"
        .FilterIndex = 4
        fileName = .Show
    End With
    
    If fileName <> "" Then
        wsStrat.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fd.InitialFileName, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=True
    End If
End Sub

'============================================================================
' BUTTON ASSIGNMENT CODE (Add to a separate module or ThisWorkbook)
'============================================================================

' To assign these macros to buttons, add this code to a worksheet module:
'
' Example for Control Panel buttons:
' 
' Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'     If Target.Address = "$B$14" Then ' Load Data button
'         Call LoadLoanTapeData
'     ElseIf Target.Address = "$B$15" Then ' Configure Fields
'         Sheets("Field Definitions").Activate
'     ElseIf Target.Address = "$B$16" Then ' Configure Buckets
'         Sheets("Bucket Definitions").Activate
'     ElseIf Target.Address = "$B$17" Then ' Generate Stratifications
'         Call GenerateAllStratifications
'     ElseIf Target.Address = "$B$18" Then ' Export to PDF
'         Call ExportToPDF
'     End If
' End Sub
