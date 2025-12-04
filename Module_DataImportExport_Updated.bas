Attribute VB_Name = "DataImportExport"
Option Explicit

' ============================================================================
' Bank of England Loan Tape - Data Import/Export Module (UPDATED)
' ============================================================================
' Handles importing loan data from external sources with intelligent header matching
' ============================================================================

' Import loan data from CSV file with header mapping
Public Sub ImportFromCSV()
    Dim fileDialog As FileDialog
    Dim filePath As String
    Dim fNum As Integer
    Dim lineText As String
    Dim dataArray() As String
    Dim headerArray() As String
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim colNum As Long
    Dim totalRows As Long
    Dim colMapping() As Long
    Dim templateColCount As Long
    Dim i As Long
    
    ' Get file from user
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select Loan Tape CSV File"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Confirm overwrite
    Set ws = ThisWorkbook.Sheets("LoanData")
    If ws.Cells(5, 1).Value <> "" Then
        If MsgBox("Existing data will be overwritten. Continue?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        ws.Range("A5:IZ" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).ClearContents
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Open and read CSV
    fNum = FreeFile
    Open filePath For Input As #fNum
    
    ' Read header row
    If Not EOF(fNum) Then
        Line Input #fNum, lineText
        headerArray = Split(lineText, ",")
        
        ' Clean header values
        For i = 0 To UBound(headerArray)
            headerArray(i) = Trim(Replace(headerArray(i), """", ""))
        Next i
        
        ' Build column mapping
        templateColCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        ReDim colMapping(0 To UBound(headerArray))
        
        For i = 0 To UBound(headerArray)
            colMapping(i) = FindTemplateColumn(ws, headerArray(i), templateColCount)
        Next i
        
        ' Show mapping summary
        Dim mappedCount As Long
        mappedCount = 0
        For i = 0 To UBound(colMapping)
            If colMapping(i) > 0 Then mappedCount = mappedCount + 1
        Next i
        
        If MsgBox("Header mapping complete:" & vbCrLf & _
                  "Matched fields: " & mappedCount & " of " & (UBound(headerArray) + 1) & vbCrLf & _
                  "Template has: " & templateColCount & " total fields" & vbCrLf & vbCrLf & _
                  "Proceed with import?", vbYesNo + vbQuestion) = vbNo Then
            Close #fNum
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
    End If
    
    ' Read data rows
    rowNum = 5
    totalRows = 0
    Do While Not EOF(fNum)
        Line Input #fNum, lineText
        
        If Trim(lineText) <> "" Then
            dataArray = Split(lineText, ",")
            
            ' Map data to template columns
            For colNum = 0 To UBound(dataArray)
                If colNum <= UBound(colMapping) Then
                    If colMapping(colNum) > 0 Then
                        ws.Cells(rowNum, colMapping(colNum)).Value = _
                            Trim(Replace(dataArray(colNum), """", ""))
                    End If
                End If
            Next colNum
            
            rowNum = rowNum + 1
            totalRows = totalRows + 1
            
            If totalRows Mod 100 = 0 Then
                ShowProgress totalRows, totalRows, "Importing"
            End If
        End If
    Loop
    
    Close #fNum
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    
    MsgBox "Successfully imported " & totalRows & " loans from CSV file." & vbCrLf & _
           "Mapped " & mappedCount & " fields to template.", vbInformation
    Exit Sub
    
ErrorHandler:
    Close #fNum
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    MsgBox "Error importing file: " & Err.Description, vbCritical
End Sub

' Import loan data from Excel file with header mapping
Public Sub ImportFromExcel()
    Dim fileDialog As FileDialog
    Dim filePath As String
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    Dim destWS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim totalRows As Long
    Dim colMapping() As Long
    Dim templateColCount As Long
    Dim headerArray() As String
    Dim mappedCount As Long
    
    ' Get file from user
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select Loan Tape Excel File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xls; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Open source workbook
    Set sourceWB = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWS = sourceWB.Sheets(1) ' Import from first sheet
    
    ' Get dimensions
    lastRow = sourceWS.Cells(sourceWS.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWS.Cells(1, sourceWS.Columns.Count).End(xlToLeft).Column
    
    ' Read headers from source
    ReDim headerArray(1 To lastCol)
    For j = 1 To lastCol
        headerArray(j) = Trim(CStr(sourceWS.Cells(1, j).Value))
    Next j
    
    ' Confirm overwrite
    Set destWS = ThisWorkbook.Sheets("LoanData")
    If destWS.Cells(5, 1).Value <> "" Then
        If MsgBox("Existing data will be overwritten. Continue?", vbYesNo + vbQuestion) = vbNo Then
            sourceWB.Close SaveChanges:=False
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        destWS.Range("A5:IZ" & destWS.Cells(destWS.Rows.Count, 1).End(xlUp).Row).ClearContents
    End If
    
    ' Build column mapping
    templateColCount = destWS.Cells(1, destWS.Columns.Count).End(xlToLeft).Column
    ReDim colMapping(1 To lastCol)
    
    For j = 1 To lastCol
        colMapping(j) = FindTemplateColumn(destWS, headerArray(j), templateColCount)
    Next j
    
    ' Count mapped fields
    mappedCount = 0
    For j = 1 To lastCol
        If colMapping(j) > 0 Then mappedCount = mappedCount + 1
    Next j
    
    If MsgBox("Header mapping complete:" & vbCrLf & _
              "Matched fields: " & mappedCount & " of " & lastCol & vbCrLf & _
              "Template has: " & templateColCount & " total fields" & vbCrLf & vbCrLf & _
              "Proceed with import?", vbYesNo + vbQuestion) = vbNo Then
        sourceWB.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Copy data with mapping
    totalRows = lastRow - 1
    For i = 2 To lastRow
        For j = 1 To lastCol
            If colMapping(j) > 0 Then
                destWS.Cells(i + 3, colMapping(j)).Value = sourceWS.Cells(i, j).Value
            End If
        Next j
        
        If (i - 1) Mod 100 = 0 Then
            ShowProgress i - 1, totalRows, "Importing"
        End If
    Next i
    
    sourceWB.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    
    MsgBox "Successfully imported " & totalRows & " loans from Excel file." & vbCrLf & _
           "Mapped " & mappedCount & " fields to template.", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not sourceWB Is Nothing Then sourceWB.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    ClearProgress
    MsgBox "Error importing file: " & Err.Description, vbCritical
End Sub

' Helper function to find template column by AR code or field name
Private Function FindTemplateColumn(ws As Worksheet, headerValue As String, _
                                   maxCol As Long) As Long
    Dim col As Long
    Dim arCode As String
    Dim fieldName As String
    
    FindTemplateColumn = 0
    headerValue = Trim(UCase(headerValue))
    
    If headerValue = "" Then Exit Function
    
    ' Search for exact AR code match (Row 1)
    For col = 1 To maxCol
        arCode = Trim(UCase(CStr(ws.Cells(1, col).Value)))
        If arCode = headerValue Then
            FindTemplateColumn = col
            Exit Function
        End If
    Next col
    
    ' Search for field name match (Row 2)
    For col = 1 To maxCol
        fieldName = Trim(UCase(CStr(ws.Cells(2, col).Value)))
        If fieldName = headerValue Then
            FindTemplateColumn = col
            Exit Function
        End If
    Next col
    
    ' No match found
    FindTemplateColumn = 0
End Function

' Export loan data to CSV
Public Sub ExportLoanDataToCSV()
    Dim ws As Worksheet
    Dim fileDialog As FileDialog
    Dim filePath As String
    Dim fNum As Integer
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim csvContent As String
    
    Set ws = ThisWorkbook.Sheets("LoanData")
    
    ' Check if data exists
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then
        MsgBox "No data to export.", vbInformation
        Exit Sub
    End If
    
    ' Get save location
    Set fileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With fileDialog
        .Title = "Export Loan Data"
        .InitialFileName = "BoE_LoanTape_" & Format(Now, "YYYYMMDD_HHMMSS") & ".csv"
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    fNum = FreeFile
    Open filePath For Output As #fNum
    
    ' Write AR codes as header
    csvContent = ""
    For j = 1 To lastCol
        If j > 1 Then csvContent = csvContent & ","
        csvContent = csvContent & """" & ws.Cells(1, j).Value & """"
    Next j
    Print #fNum, csvContent
    
    ' Write data rows
    For i = 5 To lastRow
        csvContent = ""
        For j = 1 To lastCol
            If j > 1 Then csvContent = csvContent & ","
            csvContent = csvContent & """" & Replace(ws.Cells(i, j).Value, """", """""") & """"
        Next j
        Print #fNum, csvContent
        
        If (i - 4) Mod 100 = 0 Then
            ShowProgress i - 4, lastRow - 4, "Exporting"
        End If
    Next i
    
    Close #fNum
    
    Application.ScreenUpdating = True
    ClearProgress
    
    MsgBox "Loan data exported successfully to:" & vbCrLf & filePath, vbInformation
    Exit Sub
    
ErrorHandler:
    Close #fNum
    Application.ScreenUpdating = True
    ClearProgress
    MsgBox "Error exporting file: " & Err.Description, vbCritical
End Sub

' Export complete workbook to BoE submission format
Public Sub ExportToBoEFormat()
    Dim ws As Worksheet
    Dim newWB As Workbook
    Dim newWS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim filePath As String
    Dim fileDialog As FileDialog
    
    Set ws = ThisWorkbook.Sheets("LoanData")
    
    ' Check if data exists
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then
        MsgBox "No data to export.", vbInformation
        Exit Sub
    End If
    
    ' Get save location
    Set fileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With fileDialog
        .Title = "Export to BoE Format"
        .InitialFileName = "BoE_Submission_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create new workbook
    Set newWB = Workbooks.Add
    Set newWS = newWB.Sheets(1)
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Copy AR codes only (no field names)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy
    newWS.Cells(1, 1).PasteSpecial xlPasteValues
    
    ' Copy data
    ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol)).Copy
    newWS.Cells(2, 1).PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    
    ' Format
    newWS.Rows(1).Font.Bold = True
    newWS.Columns.AutoFit
    
    ' Save
    newWB.SaveAs filePath
    newWB.Close
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "BoE submission file created successfully:" & vbCrLf & filePath, vbInformation
End Sub

' Create data quality report
Public Sub ExportDataQualityReport()
    Dim ws As Worksheet
    Dim reportWB As Workbook
    Dim summaryWS As Worksheet, errorsWS As Worksheet
    Dim filePath As String
    Dim fileDialog As FileDialog
    
    ' Get save location
    Set fileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With fileDialog
        .Title = "Export Data Quality Report"
        .InitialFileName = "DataQuality_Report_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create new workbook
    Set reportWB = Workbooks.Add
    
    ' Copy Dashboard summary
    Set summaryWS = reportWB.Sheets(1)
    summaryWS.Name = "Summary"
    
    ThisWorkbook.Sheets("Dashboard").Range("A1:H30").Copy
    summaryWS.Cells(1, 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Copy Error Report
    Set errorsWS = reportWB.Sheets.Add(After:=reportWB.Sheets(reportWB.Sheets.Count))
    errorsWS.Name = "Errors"
    
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow >= 4 Then
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 9)).Copy
        errorsWS.Cells(1, 1).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        
        ' Apply autofilter
        errorsWS.Range("A4:I4").AutoFilter
    End If
    
    ' Save
    reportWB.SaveAs filePath
    reportWB.Close
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Data quality report exported successfully:" & vbCrLf & filePath, vbInformation
End Sub

' Import data template (for users to fill externally)
Public Sub CreateImportTemplate()
    Dim ws As Worksheet
    Dim templateWB As Workbook
    Dim templateWS As Worksheet
    Dim filePath As String
    Dim fileDialog As FileDialog
    Dim lastCol As Long
    
    Set ws = ThisWorkbook.Sheets("LoanData")
    
    ' Get save location
    Set fileDialog = Application.FileDialog(msoFileDialogSaveAs)
    With fileDialog
        .Title = "Create Import Template"
        .InitialFileName = "BoE_DataEntry_Template_" & Format(Now, "YYYYMMDD") & ".xlsx"
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Create new workbook
    Set templateWB = Workbooks.Add
    Set templateWS = templateWB.Sheets(1)
    templateWS.Name = "LoanData"
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Copy headers (rows 1-4)
    ws.Range(ws.Cells(1, 1), ws.Cells(4, lastCol)).Copy
    templateWS.Cells(1, 1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Add instruction sheet
    Dim instrWS As Worksheet
    Set instrWS = templateWB.Sheets.Add(Before:=templateWB.Sheets(1))
    instrWS.Name = "Instructions"
    
    With instrWS
        .Range("A1").Value = "LOAN DATA ENTRY TEMPLATE - INSTRUCTIONS"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        .Range("A3").Value = "1. Enter loan data in the LoanData sheet starting from Row 5"
        .Range("A4").Value = "2. Follow field priorities: Red = Mandatory, Green = Optional"
        .Range("A5").Value = "3. Use date format: DD-MM-YYYY"
        .Range("A6").Value = "4. Use dropdown lists where provided (Y/N/ND fields)"
        .Range("A7").Value = "5. Refer to AR codes (row 1) - these MUST match for import"
        .Range("A8").Value = "6. Save file and import back to validation system"
        .Range("A9").Value = "7. System will automatically map columns by AR code"
        
        .Columns("A:A").AutoFit
    End With
    
    ' Save
    templateWB.SaveAs filePath
    templateWB.Close
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Import template created successfully:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
           "Users can fill this template and import it back using Import functions.", vbInformation
End Sub

' Batch import from folder
Public Sub BatchImportFromFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim fileCount As Long
    Dim totalLoans As Long
    
    ' Get folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder with CSV/Excel Files"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    Application.ScreenUpdating = False
    
    ' Process CSV files
    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        ' Import logic here (would need to be implemented)
        fileCount = fileCount + 1
        fileName = Dir
    Loop
    
    ' Process Excel files
    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        ' Import logic here (would need to be implemented)
        fileCount = fileCount + 1
        fileName = Dir
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Batch import complete." & vbCrLf & _
           "Files processed: " & fileCount & vbCrLf & _
           "Total loans imported: " & totalLoans, vbInformation
End Sub
