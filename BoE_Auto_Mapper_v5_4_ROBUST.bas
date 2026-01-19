Attribute VB_Name = "BoE_Auto_Mapper"
Option Explicit

' ===================================================================
' BOE AUTO-MAPPER - ROBUST ERROR HANDLING
' Version 5.4 - Production-Grade Data Validation
' ===================================================================

' Module-level error log
Private errorLog As Collection

Sub AddMapperButton()
    Dim ws As Worksheet
    Dim btn As Button
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Loan Tape (BoE Raw)")
    
    If ws Is Nothing Then
        MsgBox "Sheet 'Loan Tape (BoE Raw)' not found!" & vbCrLf & vbCrLf & _
               "Please ensure the model structure is intact.", vbCritical, _
               "Sheet Not Found"
        Exit Sub
    End If
    
    ws.Buttons.Delete
    
    Set btn = ws.Buttons.Add(10, 10, 200, 40)
    btn.OnAction = "MapBoEFields"
    btn.Text = "Map BoE Fields" & vbCrLf & "(AR Codes)"
    btn.Font.Bold = True
    btn.Font.Size = 11
    btn.Font.Color = RGB(255, 255, 255)
    
    With btn
        .ShapeRange.Fill.ForeColor.RGB = RGB(54, 96, 146)
        .ShapeRange.Line.ForeColor.RGB = RGB(54, 96, 146)
    End With
    
    MsgBox "✓ Button Added Successfully!" & vbCrLf & vbCrLf & _
           "Model Version: 5.4 - Robust Error Handling" & vbCrLf & vbCrLf & _
           "Features:" & vbCrLf & _
           "• 26 List-type AR fields with code conversions" & vbCrLf & _
           "• Robust date/numeric validation" & vbCrLf & _
           "• Comprehensive error logging" & vbCrLf & _
           "• Format enforcement and recovery" & vbCrLf & vbCrLf & _
           "Paste BoE PRA110 loan tape and click 'Map BoE Fields'", _
           vbInformation, "BoE Auto-Mapper v5.4"
End Sub

Sub MapBoEFields()
    Dim wsRaw As Worksheet, wsMapper As Worksheet, wsBoE As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim arCode As String, boEColumn As Long, conversionRule As String
    Dim colMapping As Object, headerRow As Long
    Dim startTime As Double
    Dim mappedFields As Long, mappedLoans As Long
    Dim criticalMissing As Object
    
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Initialize error log
    Set errorLog = New Collection
    
    Set wsRaw = ThisWorkbook.Sheets("Loan Tape (BoE Raw)")
    Set wsMapper = ThisWorkbook.Sheets("BoE Auto-Mapper")
    Set wsBoE = ThisWorkbook.Sheets("Loan Tape (BoE)")
    
    headerRow = FindARCodeRow(wsRaw)
    If headerRow = 0 Then
        MsgBox "❌ AR Codes Not Found", vbCritical
        GoTo CleanUp
    End If
    
    Set colMapping = BuildMappingDictionary(wsMapper)
    If colMapping.Count = 0 Then
        MsgBox "❌ No Field Mappings", vbCritical
        GoTo CleanUp
    End If
    
    lastCol = wsRaw.Cells(headerRow, wsRaw.Columns.Count).End(xlToLeft).Column
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= headerRow Then
        MsgBox "❌ No Loan Data Found", vbExclamation
        GoTo CleanUp
    End If
    
    mappedLoans = lastRow - headerRow
    
    If mappedLoans > 10000 Then
        MsgBox "Processing " & Format(mappedLoans, "#,##0") & " loans..." & vbCrLf & _
               "Please wait", vbInformation, "Large File"
    End If
    
    wsBoE.Range("A5:AZ" & wsBoE.Cells(wsBoE.Rows.Count, 1).End(xlUp).Row).ClearContents
    
    Set criticalMissing = CreateObject("Scripting.Dictionary")
    InitializeCriticalFields criticalMissing
    
    mappedFields = 0
    
    ' Map fields with error handling
    For j = 1 To lastCol
        Dim headerText As String, extractedCode As String
        headerText = UCase(Trim(wsRaw.Cells(headerRow, j).Value))
        extractedCode = ExtractARCode(headerText)
        
        If extractedCode <> "" And colMapping.Exists(extractedCode) Then
            boEColumn = colMapping(extractedCode)(0)
            conversionRule = colMapping(extractedCode)(1)
            
            For i = headerRow + 1 To lastRow
                On Error Resume Next
                Dim rawValue As Variant, convertedValue As Variant
                rawValue = wsRaw.Cells(i, j).Value
                
                ' Convert with error handling
                convertedValue = ConvertValueSafe(rawValue, conversionRule, extractedCode, i - headerRow)
                
                If Err.Number <> 0 Then
                    LogError extractedCode, i - headerRow, rawValue, Err.Description
                    Err.Clear
                End If
                
                wsBoE.Cells(i - headerRow + 4, boEColumn).Value = convertedValue
                On Error GoTo ErrorHandler
            Next i
            
            mappedFields = mappedFields + 1
            
            If criticalMissing.Exists(extractedCode) Then
                criticalMissing.Remove extractedCode
            End If
        End If
    Next j
    
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    Dim reportMsg As String
    reportMsg = GenerateDataQualityReport(mappedFields, mappedLoans, _
                                         elapsedTime, colMapping.Count, _
                                         criticalMissing)
    
    ' Add error summary if errors occurred
    If errorLog.Count > 0 Then
        reportMsg = reportMsg & vbCrLf & vbCrLf & _
                   "⚠️ DATA QUALITY ISSUES: " & errorLog.Count & vbCrLf & _
                   "Run 'ViewErrorLog' to see details"
    End If
    
    If criticalMissing.Count = 0 And errorLog.Count = 0 Then
        MsgBox reportMsg, vbInformation, "Mapping Complete"
    ElseIf errorLog.Count > 0 Then
        MsgBox reportMsg, vbExclamation, "Mapping Complete - Data Issues"
    Else
        MsgBox reportMsg, vbCritical, "Mapping Complete - Critical Issues"
    End If
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ ERROR: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub

Function ConvertValueSafe(rawValue As Variant, conversionRule As String, _
                         arCode As String, rowNum As Long) As Variant
    ' Safe wrapper around ConvertValue with error handling
    
    On Error Resume Next
    
    If IsEmpty(rawValue) Or rawValue = "" Or IsNull(rawValue) Then
        ConvertValueSafe = ""
        Exit Function
    End If
    
    ' Determine data type and apply appropriate validation
    Select Case True
        Case InStr(conversionRule, "Date") > 0
            ConvertValueSafe = ParseDateSafe(rawValue, arCode, rowNum)
            
        Case InStr(conversionRule, "Number") > 0 Or InStr(conversionRule, "Percentage") > 0
            ConvertValueSafe = ParseNumericSafe(rawValue, arCode, rowNum)
            
        Case conversionRule = "YN_Convert" Or conversionRule = "Code_Lookup"
            ConvertValueSafe = ConvertCodeSafe(rawValue, arCode, rowNum)
            
        Case conversionRule = "Direct" Or conversionRule = ""
            ConvertValueSafe = rawValue
            
        Case Else
            ConvertValueSafe = rawValue
    End Select
    
    If Err.Number <> 0 Then
        LogError arCode, rowNum, rawValue, "Conversion failed: " & Err.Description
        ConvertValueSafe = ""
        Err.Clear
    End If
End Function

Function ParseDateSafe(rawValue As Variant, arCode As String, rowNum As Long) As Variant
    ' Robust date parsing with multiple format support
    
    On Error Resume Next
    Dim testVal As Variant, dateVal As Date
    
    ' Handle common invalid values
    Select Case UCase(Trim(CStr(rawValue)))
        Case "N/A", "NA", "TBC", "TBD", "-", "NULL", "NONE"
            ParseDateSafe = ""
            Exit Function
        Case "ND", "NO DATA"
            ParseDateSafe = ""
            Exit Function
    End Select
    
    ' Method 1: Try direct IsDate
    If IsDate(rawValue) Then
        ParseDateSafe = CDate(rawValue)
        Exit Function
    End If
    
    ' Method 2: Try Excel serial number
    If IsNumeric(rawValue) Then
        testVal = CDbl(rawValue)
        If testVal > 0 And testVal < 2958466 Then ' Valid Excel date range
            On Error Resume Next
            dateVal = CDate(testVal)
            If Err.Number = 0 Then
                ParseDateSafe = dateVal
                Exit Function
            End If
            Err.Clear
        End If
    End If
    
    ' Method 3: Try parsing text dates
    Dim strDate As String
    strDate = Trim(CStr(rawValue))
    
    ' Try DD/MM/YYYY
    If Len(strDate) = 10 And Mid(strDate, 3, 1) = "/" And Mid(strDate, 6, 1) = "/" Then
        On Error Resume Next
        dateVal = DateSerial(CLng(Right(strDate, 4)), CLng(Mid(strDate, 4, 2)), CLng(Left(strDate, 2)))
        If Err.Number = 0 Then
            ParseDateSafe = dateVal
            Exit Function
        End If
        Err.Clear
    End If
    
    ' Try YYYY-MM-DD
    If Len(strDate) = 10 And Mid(strDate, 5, 1) = "-" And Mid(strDate, 8, 1) = "-" Then
        On Error Resume Next
        dateVal = DateSerial(CLng(Left(strDate, 4)), CLng(Mid(strDate, 6, 2)), CLng(Right(strDate, 2)))
        If Err.Number = 0 Then
            ParseDateSafe = dateVal
            Exit Function
        End If
        Err.Clear
    End If
    
    ' Method 4: Try VBA DateValue
    On Error Resume Next
    dateVal = DateValue(strDate)
    If Err.Number = 0 Then
        ParseDateSafe = dateVal
        Exit Function
    End If
    Err.Clear
    
    ' All methods failed
    LogError arCode, rowNum, rawValue, "Invalid date format"
    ParseDateSafe = ""
End Function

Function ParseNumericSafe(rawValue As Variant, arCode As String, rowNum As Long) As Variant
    ' Robust numeric parsing with format cleanup
    
    On Error Resume Next
    Dim cleanVal As String, numVal As Double
    
    ' Handle common invalid values
    Select Case UCase(Trim(CStr(rawValue)))
        Case "N/A", "NA", "TBC", "TBD", "-", "NULL", "NONE", "#N/A", "#DIV/0!"
            ParseNumericSafe = 0
            Exit Function
        Case "ND", "NO DATA"
            ParseNumericSafe = 0
            Exit Function
    End Select
    
    ' If already numeric, return it
    If IsNumeric(rawValue) Then
        ParseNumericSafe = CDbl(rawValue)
        Exit Function
    End If
    
    ' Clean text-formatted numbers
    cleanVal = Trim(CStr(rawValue))
    
    ' Remove currency symbols
    cleanVal = Replace(cleanVal, "£", "")
    cleanVal = Replace(cleanVal, "$", "")
    cleanVal = Replace(cleanVal, "€", "")
    
    ' Handle percentages
    Dim isPercentage As Boolean
    isPercentage = (Right(cleanVal, 1) = "%")
    If isPercentage Then
        cleanVal = Left(cleanVal, Len(cleanVal) - 1)
    End If
    
    ' Remove thousand separators
    cleanVal = Replace(cleanVal, ",", "")
    
    ' Remove spaces
    cleanVal = Replace(cleanVal, " ", "")
    
    ' Try conversion
    If IsNumeric(cleanVal) Then
        numVal = CDbl(cleanVal)
        
        ' If was percentage, convert to decimal
        If isPercentage Then
            numVal = numVal / 100
        End If
        
        ParseNumericSafe = numVal
        Exit Function
    End If
    
    ' All methods failed
    LogError arCode, rowNum, rawValue, "Invalid numeric format"
    ParseNumericSafe = 0
End Function

Function ConvertCodeSafe(rawValue As Variant, arCode As String, rowNum As Long) As Variant
    ' Safe code conversion with validation
    
    On Error Resume Next
    Dim code As String
    code = UCase(Trim(CStr(rawValue)))
    
    ' Handle Y/N/ND fields specially
    If arCode = "AR22" Or arCode = "AR23" Or arCode = "AR36" Or arCode = "AR47" Then
        Select Case code
            Case "Y", "YES", "1", "TRUE"
                ConvertCodeSafe = "Yes"
            Case "N", "NO", "0", "FALSE", ""
                ConvertCodeSafe = "No"
            Case "ND", "NO DATA", "N/A", "NA"
                ConvertCodeSafe = "No Data"
            Case Else
                LogError arCode, rowNum, rawValue, "Invalid Y/N/ND value, defaulting to No"
                ConvertCodeSafe = "No"
        End Select
        Exit Function
    End If
    
    ' For other codes, use main conversion logic
    ConvertCodeSafe = ConvertValue(rawValue, "", arCode)
    
    ' Validate the conversion worked
    If ConvertCodeSafe = rawValue And IsNumeric(rawValue) Then
        ' Code wasn't converted - may be invalid
        If Not IsValidCode(arCode, code) Then
            LogError arCode, rowNum, rawValue, "Unrecognized code - passing through"
        End If
    End If
End Function

Function ConvertValue(rawValue As Variant, conversionRule As String, arCode As String) As Variant
    ' Core conversion logic (preserved from v5.3)
    
    If IsEmpty(rawValue) Or rawValue = "" Or IsNull(rawValue) Then
        ConvertValue = ""
        Exit Function
    End If
    
    Dim code As String
    code = UCase(Trim(CStr(rawValue)))
    
    Select Case arCode
        Case "AR21", "AR189"
            Select Case code
                Case "1": ConvertValue = "Employed"
                Case "2": ConvertValue = "Self-Employed"
                Case "3": ConvertValue = "Retired"
                Case "4": ConvertValue = "Not Employed"
                Case "5": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR22", "AR23"
            Select Case code
                Case "Y", "YES": ConvertValue = "Yes"
                Case "N", "NO": ConvertValue = "No"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = "No"
            End Select
            
        Case "AR27", "AR29"
            Select Case code
                Case "1": ConvertValue = "Self-Certified"
                Case "2": ConvertValue = "Self-Cert with Affordability"
                Case "3": ConvertValue = "Verified"
                Case "4": ConvertValue = "Non-Verified"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR58"
            Select Case code
                Case "1": ConvertValue = "Branch Network"
                Case "2": ConvertValue = "Direct"
                Case "3": ConvertValue = "Broker"
                Case "4": ConvertValue = "Internet"
                Case "5": ConvertValue = "Packager"
                Case "6": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR59"
            Select Case code
                Case "1": ConvertValue = "Purchase"
                Case "2": ConvertValue = "Remortgage"
                Case "3": ConvertValue = "Renovation"
                Case "4": ConvertValue = "Equity Release"
                Case "5": ConvertValue = "Construction"
                Case "6": ConvertValue = "Debt Consolidation"
                Case "7": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR60"
            Select Case code
                Case "1": ConvertValue = "No"
                Case "2", "3", "4", "5", "6": ConvertValue = "Yes"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = "No"
            End Select
            
        Case "AR69"
            Select Case code
                Case "1": ConvertValue = "Interest Only"
                Case "2": ConvertValue = "Repayment"
                Case "3": ConvertValue = "Endowment"
                Case "4": ConvertValue = "Pension"
                Case "5": ConvertValue = "ISA/PEP"
                Case "6": ConvertValue = "Index-Linked"
                Case "7": ConvertValue = "Part and Part"
                Case "8": ConvertValue = "Savings Mortgage"
                Case "9": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        Case "AR70"
            Select Case code
                Case "1": ConvertValue = "Monthly"
                Case "2": ConvertValue = "Quarterly"
                Case "3": ConvertValue = "Semi-Annually"
                Case "4": ConvertValue = "Annual"
                Case "5": ConvertValue = "Bullet"
                Case "6": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR72"
            Select Case code
                Case "1": ConvertValue = "Annuity"
                Case "2": ConvertValue = "Linear"
                Case "3": ConvertValue = "Increasing Instalments"
                Case "4": ConvertValue = "Fixed (Changing Maturity)"
                Case "5": ConvertValue = "Fixed (Changing Interest)"
                Case "6": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR74"
            Select Case code
                Case "1": ConvertValue = "No Guarantor"
                Case "2": ConvertValue = "Individual - Family"
                Case "3": ConvertValue = "Individual - Other"
                Case "4": ConvertValue = "Government"
                Case "5": ConvertValue = "Bank"
                Case "6": ConvertValue = "Insurance"
                Case "7": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR84"
            Select Case code
                Case "1": ConvertValue = "First"
                Case "2": ConvertValue = "Second"
                Case "3": ConvertValue = "Third"
                Case "4": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        Case "AR105"
            Select Case code
                Case "1", "2", "3": ConvertValue = "Yes"
                Case "4": ConvertValue = "No"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = "No"
            End Select
            
        Case "AR107"
            Select Case code
                Case "1", "6": ConvertValue = "Variable"
                Case "2": ConvertValue = "Tracker"
                Case "3", "4", "5": ConvertValue = "Fixed"
                Case "7": ConvertValue = "Discount"
                Case "8": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        Case "AR108"
            Select Case code
                Case "1": ConvertValue = "1M LIBOR"
                Case "2": ConvertValue = "1M EURIBOR"
                Case "3": ConvertValue = "3M LIBOR"
                Case "4": ConvertValue = "3M EURIBOR"
                Case "5": ConvertValue = "6M LIBOR"
                Case "6": ConvertValue = "6M EURIBOR"
                Case "7": ConvertValue = "12M LIBOR"
                Case "8": ConvertValue = "12M EURIBOR"
                Case "9": ConvertValue = "BoE Base Rate"
                Case "10": ConvertValue = "ECB Base Rate"
                Case "11": ConvertValue = "Lender SVR"
                Case "12": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR123"
            Select Case code
                Case "1": ConvertValue = "Capitalization"
                Case "2": ConvertValue = "Payment Holiday"
                Case "3": ConvertValue = "Term Extension"
                Case "4": ConvertValue = "Interest Rate Concession"
                Case "5": ConvertValue = "Interest Only Period"
                Case "6": ConvertValue = "Combination"
                Case "7": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR128"
            Select Case code
                Case "1": ConvertValue = "England - London"
                Case "2": ConvertValue = "England - South East"
                Case "3": ConvertValue = "England - South West"
                Case "4": ConvertValue = "England - East Anglia"
                Case "5": ConvertValue = "England - East Midlands"
                Case "6": ConvertValue = "England - West Midlands"
                Case "7": ConvertValue = "England - Yorkshire & Humberside"
                Case "8": ConvertValue = "England - North West"
                Case "9": ConvertValue = "England - North"
                Case "10": ConvertValue = "Wales"
                Case "11": ConvertValue = "Scotland"
                Case "12": ConvertValue = "Northern Ireland"
                Case "13": ConvertValue = "Isle of Man/Channel Islands"
                Case "14": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR130"
            Select Case code
                Case "1": ConvertValue = "Owner Occupied"
                Case "2": ConvertValue = "Partially Owner Occupied"
                Case "3": ConvertValue = "Buy to Let"
                Case "4": ConvertValue = "Second Home"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        Case "AR131"
            Select Case code
                Case "1": ConvertValue = "House"
                Case "2": ConvertValue = "Flat"
                Case "3": ConvertValue = "Bungalow"
                Case "4": ConvertValue = "Terraced House"
                Case "5": ConvertValue = "Multifamily (with recourse)"
                Case "6": ConvertValue = "Multifamily (without recourse)"
                Case "7": ConvertValue = "Partially Commercial"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        Case "AR137", "AR144"
            Select Case code
                Case "1": ConvertValue = "Full"
                Case "2": ConvertValue = "External Only"
                Case "3": ConvertValue = "Drive-by"
                Case "4": ConvertValue = "AVM"
                Case "5": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR166"
            Select Case code
                Case "1": ConvertValue = "Performing"
                Case "2": ConvertValue = "Arrears"
                Case "3": ConvertValue = "Default/Foreclosure"
                Case "4": ConvertValue = "Redeemed"
                Case "5": ConvertValue = "Repurchased"
                Case "6": ConvertValue = "Other"
                Case "ND", "NO DATA": ConvertValue = "No Data"
                Case Else: ConvertValue = code
            End Select
            
        Case "AR173"
            If IsDate(rawValue) Then
                ConvertValue = "Yes"
            ElseIf code = "ND" Or code = "NO DATA" Then
                ConvertValue = "No Data"
            Else
                ConvertValue = "No"
            End If
            
        Case Else
            ConvertValue = rawValue
    End Select
End Function

Function IsValidCode(arCode As String, code As String) As Boolean
    ' Validate if a code is recognized for the given AR field
    
    IsValidCode = True  ' Assume valid by default
    
    ' Define valid codes for each field
    Select Case arCode
        Case "AR21", "AR189"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "5" Or code = "ND")
        Case "AR22", "AR23"
            IsValidCode = (code = "Y" Or code = "N" Or code = "ND")
        Case "AR27", "AR29"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "ND")
        Case "AR59"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "5" Or code = "6" Or code = "7" Or code = "ND")
        Case "AR60"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "5" Or code = "6" Or code = "ND")
        Case "AR84"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "ND")
        Case "AR105"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "ND")
        Case "AR130"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "ND" Or code = "")
        ' Add more as needed
    End Select
End Function

Sub LogError(arCode As String, rowNum As Long, rawValue As Variant, errorMsg As String)
    ' Log data quality issues for review
    
    On Error Resume Next
    Dim errorEntry As String
    errorEntry = "Row " & rowNum & " | " & arCode & " | " & _
                 "Value: '" & rawValue & "' | " & errorMsg
    errorLog.Add errorEntry
End Sub

Sub ViewErrorLog()
    ' Display error log in a message box or worksheet
    
    If errorLog Is Nothing Then
        MsgBox "No error log available. Run mapping first.", vbInformation
        Exit Sub
    End If
    
    If errorLog.Count = 0 Then
        MsgBox "✓ No errors logged", vbInformation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim wsName As String
    wsName = "Data Quality Log"
    
    ' Create or clear log sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = wsName
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Write headers
    ws.Range("A1").Value = "Row"
    ws.Range("B1").Value = "AR Code"
    ws.Range("C1").Value = "Raw Value"
    ws.Range("D1").Value = "Issue"
    ws.Range("A1:D1").Font.Bold = True
    
    ' Write errors
    Dim i As Long
    For i = 1 To errorLog.Count
        Dim parts() As String
        parts = Split(errorLog(i), " | ")
        If UBound(parts) >= 3 Then
            ws.Cells(i + 1, 1).Value = Replace(parts(0), "Row ", "")
            ws.Cells(i + 1, 2).Value = parts(1)
            ws.Cells(i + 1, 3).Value = Replace(parts(2), "Value: '", "")
            ws.Cells(i + 1, 3).Value = Replace(ws.Cells(i + 1, 3).Value, "'", "")
            ws.Cells(i + 1, 4).Value = parts(3)
        End If
    Next i
    
    ws.Columns("A:D").AutoFit
    ws.Activate
    
    MsgBox "Error log created in sheet '" & wsName & "'" & vbCrLf & _
           "Total issues: " & errorLog.Count, vbInformation, "Data Quality Log"
End Sub

' ===================================================================
' HELPER FUNCTIONS (Preserved from v5.3)
' ===================================================================

Function FindARCodeRow(ws As Worksheet) As Long
    Dim i As Long
    For i = 1 To 20
        If InStr(1, ws.Cells(i, 1).Value, "AR", vbTextCompare) > 0 Then
            FindARCodeRow = i
            Exit Function
        End If
    Next i
    FindARCodeRow = 0
End Function

Function BuildMappingDictionary(wsMapper As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, arCode As String, boEColumn As Long, conversionRule As String
    Dim mapperLastRow As Long
    mapperLastRow = wsMapper.Cells(wsMapper.Rows.Count, 1).End(xlUp).Row
    
    For i = 5 To mapperLastRow
        arCode = Trim(wsMapper.Cells(i, 1).Value)
        If arCode <> "" And Left(arCode, 2) = "AR" Then
            boEColumn = ColumnLetterToNumber(wsMapper.Cells(i, 3).Value)
            conversionRule = wsMapper.Cells(i, 6).Value
            
            If boEColumn > 0 Then
                dict(UCase(arCode)) = Array(boEColumn, conversionRule)
            End If
        End If
    Next i
    
    Set BuildMappingDictionary = dict
End Function

Sub InitializeCriticalFields(dict As Object)
    dict("AR3") = "Loan Identifier|MANDATORY"
    dict("AR66") = "Original Balance|MANDATORY"
    dict("AR67") = "Current Balance|MANDATORY"
    dict("AR136") = "Original Property Value|MANDATORY"
    dict("AR143") = "Current Property Value|MANDATORY"
    dict("AR55") = "Origination Date|MANDATORY"
    dict("AR5") = "Originator|Adj: Originator"
    dict("AR131") = "Property Type|Validation"
    dict("AR130") = "Occupancy Type|Adj: BTL/1.7x if missing"
    dict("AR141") = "Current LTV|Adj: LTV Curve"
    dict("AR84") = "Lien Position|Adj: 1.50x Second Lien"
    dict("AR59") = "Loan Purpose|Adj: Purpose"
    dict("AR69") = "Payment Method|Adj: Interest Only"
    dict("AR56") = "Maturity Date|Calc: Remaining Term"
    dict("AR26") = "Primary Income|Adj: Income Multiple"
    dict("AR28") = "Secondary Income|Adj: Income Multiple"
    dict("AR73") = "Debt to Income|Alt: Income Multiple"
    dict("AR19") = "Number of Debtors|Logic: Joint Thresholds"
    dict("AR22") = "First Time Buyer|Adj: FTB"
    dict("AR105") = "Help to Buy|Adj: 1.35x - CRITICAL"
    dict("AR60") = "Shared Ownership|Adj: Shared Own"
    dict("AR23") = "Right to Buy|Adj: RTB"
    dict("AR31") = "CCJ Satisfied|Adj: CCJ Tiers"
    dict("AR33") = "CCJ Unsatisfied|Adj: CCJ Tiers"
    dict("AR35") = "Last CCJ Date|Adj: CCJ Age Tiers"
    dict("AR36") = "Bankruptcy/IVA|Adj: 1.50x-3.00x"
    dict("AR47") = "Prior Repossession|Adj: Repossession"
    dict("AR166") = "Account Status|Status Check"
    dict("AR169") = "Arrears Balance|Arrears Amount"
    dict("AR170") = "Months in Arrears|Adj: 1.00x-15.00x"
    dict("AR173") = "Performance Arrangement|Forbearance Flag"
    dict("AR167") = "Date Last Current|Re-performing Check"
    dict("AR168") = "Date Last in Arrears|Re-performing Check"
    dict("AR154") = "Rental Income|Calc: DSCR"
    dict("AR156") = "DSCR|Adj: BTL/DSCR"
    dict("AR21") = "Employment Status|Adj: Self-Employed"
    dict("AR27") = "Income Verification|Adj: Self-Cert"
    dict("AR107") = "Interest Rate Type|Adj: Payment Shock"
End Sub

Function GenerateDataQualityReport(mappedFields As Long, mappedLoans As Long, _
                                   elapsedTime As Double, totalConfigured As Long, _
                                   criticalMissing As Object) As String
    Dim report As String
    Dim coveragePct As Double
    
    report = "✓ MAPPING COMPLETE!" & vbCrLf & vbCrLf
    report = report & "═══════════════════════════════" & vbCrLf
    report = report & "STATISTICS" & vbCrLf
    report = report & "═══════════════════════════════" & vbCrLf
    report = report & "Fields Mapped: " & mappedFields & " of " & totalConfigured & vbCrLf
    report = report & "Loans Processed: " & Format(mappedLoans, "#,##0") & vbCrLf
    report = report & "Processing Time: " & Format(elapsedTime, "0.0") & " sec" & vbCrLf
    report = report & "Rate: " & Format(mappedLoans / elapsedTime, "#,##0") & " loans/sec" & vbCrLf
    
    coveragePct = (mappedFields / totalConfigured) * 100
    report = report & "Coverage: " & Format(coveragePct, "0.0") & "%" & vbCrLf
    
    If criticalMissing.Count = 0 Then
        report = report & vbCrLf & "✓ ALL CRITICAL FIELDS PRESENT" & vbCrLf
    Else
        report = report & vbCrLf & "⚠️ MISSING: " & criticalMissing.Count & " fields" & vbCrLf
    End If
    
    GenerateDataQualityReport = report
End Function

Function ExtractARCode(headerText As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "AR\d+"
    regex.IgnoreCase = True
    
    If regex.Test(headerText) Then
        ExtractARCode = UCase(regex.Execute(headerText)(0))
    Else
        ExtractARCode = ""
    End If
End Function

Function ColumnLetterToNumber(columnLetter As String) As Long
    On Error Resume Next
    ColumnLetterToNumber = Range(columnLetter & "1").Column
    If Err.Number <> 0 Then ColumnLetterToNumber = 0
    On Error GoTo 0
End Function

Sub ShowModelInfo()
    MsgBox "S&P WAFF/WALS MODEL" & vbCrLf & vbCrLf & _
           "Version: 5.4 - Robust Error Handling" & vbCrLf & _
           "Release: 19-Jan-2026" & vbCrLf & vbCrLf & _
           "FEATURES:" & vbCrLf & _
           "✓ 26 List-type field conversions" & vbCrLf & _
           "✓ Robust date/numeric parsing" & vbCrLf & _
           "✓ Comprehensive error logging" & vbCrLf & _
           "✓ Format enforcement" & vbCrLf & vbCrLf & _
           "Run 'ViewErrorLog' to see data quality issues", _
           vbInformation, "Model Information"
End Sub

Sub ClearMappedData()
    Dim response As VbMsgBoxResult
    response = MsgBox("Clear ALL mapped data?", vbQuestion + vbYesNo, "Clear Data")
    
    If response = vbYes Then
        Dim wsBoE As Worksheet
        Set wsBoE = ThisWorkbook.Sheets("Loan Tape (BoE)")
        wsBoE.Range("A5:AZ" & wsBoE.Cells(wsBoE.Rows.Count, 1).End(xlUp).Row).ClearContents
        MsgBox "✓ Data cleared", vbInformation
    End If
End Sub
