Attribute VB_Name = "BoE_Auto_Mapper"
Option Explicit

' ===================================================================
' BOE AUTO-MAPPER v6.0 - PRODUCTION
' Complete Restructure | 52 Sorted Fields | Robust Error Handling
' ===================================================================
'
' MAJOR CHANGES v6.0:
' - 52 fields sorted by AR code (was 44 random order)
' - 8 new fields added (AR19, AR32, AR34, AR90, AR91, AR92, AR103, AR104)
' - All conversions from BoE PRA110 template
' - Robust error handling from v5.4
' - CCJ joint calculation support
' - Shared ownership LTV support
'
' ===================================================================

Private errorLog As Collection

Sub AddMapperButton()
    Dim ws As Worksheet
    Dim btn As Button
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Loan Tape (BoE Raw)")
    
    If ws Is Nothing Then
        MsgBox "Sheet 'Loan Tape (BoE Raw)' not found!", vbCritical
        Exit Sub
    End If
    
    ws.Buttons.Delete
    
    Set btn = ws.Buttons.Add(10, 10, 200, 40)
    btn.OnAction = "MapBoEFields"
    btn.Text = "Map BoE Fields" & vbCrLf & "(v6.0 - 52 Fields)"
    btn.Font.Bold = True
    btn.Font.Size = 11
    btn.Font.Color = RGB(255, 255, 255)
    
    With btn
        .ShapeRange.Fill.ForeColor.RGB = RGB(54, 96, 146)
        .ShapeRange.Line.ForeColor.RGB = RGB(54, 96, 146)
    End With
    
    MsgBox "✓ Mapper Button Added!" & vbCrLf & vbCrLf & _
           "Version: 6.0 - Production Restructure" & vbCrLf & vbCrLf & _
           "Features:" & vbCrLf & _
           "• 52 fields sorted by AR code" & vbCrLf & _
           "• 8 new fields for CCJ/Shared Own" & vbCrLf & _
           "• Robust error handling" & vbCrLf & _
           "• CCJ joint calculation support" & vbCrLf & vbCrLf & _
           "Paste BoE PRA110 tape and click button", _
           vbInformation, "BoE Auto-Mapper v6.0"
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
    reportMsg = GenerateReport(mappedFields, mappedLoans, elapsedTime, _
                              colMapping.Count, criticalMissing)
    
    If errorLog.Count > 0 Then
        reportMsg = reportMsg & vbCrLf & vbCrLf & _
                   "⚠️ DATA ISSUES: " & errorLog.Count & vbCrLf & _
                   "Run 'ViewErrorLog' for details"
    End If
    
    If criticalMissing.Count = 0 And errorLog.Count = 0 Then
        MsgBox reportMsg, vbInformation, "Mapping Complete"
    Else
        MsgBox reportMsg, vbExclamation, "Mapping Complete - Review Issues"
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
    On Error Resume Next
    
    If IsEmpty(rawValue) Or rawValue = "" Or IsNull(rawValue) Then
        ConvertValueSafe = ""
        Exit Function
    End If
    
    Select Case True
        Case InStr(conversionRule, "Date") > 0
            ConvertValueSafe = ParseDateSafe(rawValue, arCode, rowNum)
            
        Case InStr(conversionRule, "Number") > 0 Or InStr(conversionRule, "Percentage") > 0
            ConvertValueSafe = ParseNumericSafe(rawValue, arCode, rowNum)
            
        Case conversionRule = "YN_Convert" Or conversionRule = "Code_Lookup"
            ConvertValueSafe = ConvertCodeSafe(rawValue, arCode, rowNum)
            
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
    On Error Resume Next
    Dim testVal As Variant, dateVal As Date, strDate As String
    
    Select Case UCase(Trim(CStr(rawValue)))
        Case "N/A", "NA", "TBC", "TBD", "-", "NULL", "NONE", "ND", "NO DATA"
            ParseDateSafe = ""
            Exit Function
    End Select
    
    If IsDate(rawValue) Then
        ParseDateSafe = CDate(rawValue)
        Exit Function
    End If
    
    If IsNumeric(rawValue) Then
        testVal = CDbl(rawValue)
        If testVal > 0 And testVal < 2958466 Then
            On Error Resume Next
            dateVal = CDate(testVal)
            If Err.Number = 0 Then
                ParseDateSafe = dateVal
                Exit Function
            End If
            Err.Clear
        End If
    End If
    
    strDate = Trim(CStr(rawValue))
    
    If Len(strDate) = 10 And Mid(strDate, 3, 1) = "/" And Mid(strDate, 6, 1) = "/" Then
        On Error Resume Next
        dateVal = DateSerial(CLng(Right(strDate, 4)), CLng(Mid(strDate, 4, 2)), CLng(Left(strDate, 2)))
        If Err.Number = 0 Then
            ParseDateSafe = dateVal
            Exit Function
        End If
        Err.Clear
    End If
    
    If Len(strDate) = 10 And Mid(strDate, 5, 1) = "-" And Mid(strDate, 8, 1) = "-" Then
        On Error Resume Next
        dateVal = DateSerial(CLng(Left(strDate, 4)), CLng(Mid(strDate, 6, 2)), CLng(Right(strDate, 2)))
        If Err.Number = 0 Then
            ParseDateSafe = dateVal
            Exit Function
        End If
        Err.Clear
    End If
    
    On Error Resume Next
    dateVal = DateValue(strDate)
    If Err.Number = 0 Then
        ParseDateSafe = dateVal
        Exit Function
    End If
    Err.Clear
    
    LogError arCode, rowNum, rawValue, "Invalid date format"
    ParseDateSafe = ""
End Function

Function ParseNumericSafe(rawValue As Variant, arCode As String, rowNum As Long) As Variant
    On Error Resume Next
    Dim cleanVal As String, numVal As Double, isPercentage As Boolean
    
    Select Case UCase(Trim(CStr(rawValue)))
        Case "N/A", "NA", "TBC", "TBD", "-", "NULL", "NONE", "#N/A", "#DIV/0!", "ND"
            ParseNumericSafe = 0
            Exit Function
    End Select
    
    If IsNumeric(rawValue) Then
        ParseNumericSafe = CDbl(rawValue)
        Exit Function
    End If
    
    cleanVal = Trim(CStr(rawValue))
    cleanVal = Replace(cleanVal, "£", "")
    cleanVal = Replace(cleanVal, "$", "")
    cleanVal = Replace(cleanVal, "€", "")
    
    isPercentage = (Right(cleanVal, 1) = "%")
    If isPercentage Then cleanVal = Left(cleanVal, Len(cleanVal) - 1)
    
    cleanVal = Replace(cleanVal, ",", "")
    cleanVal = Replace(cleanVal, " ", "")
    
    If IsNumeric(cleanVal) Then
        numVal = CDbl(cleanVal)
        If isPercentage Then numVal = numVal / 100
        ParseNumericSafe = numVal
        Exit Function
    End If
    
    LogError arCode, rowNum, rawValue, "Invalid numeric format"
    ParseNumericSafe = 0
End Function

Function ConvertCodeSafe(rawValue As Variant, arCode As String, rowNum As Long) As Variant
    On Error Resume Next
    Dim code As String
    code = UCase(Trim(CStr(rawValue)))
    
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
    
    ConvertCodeSafe = ConvertValue(rawValue, "", arCode)
    
    If ConvertCodeSafe = rawValue And IsNumeric(rawValue) Then
        If Not IsValidCode(arCode, code) Then
            LogError arCode, rowNum, rawValue, "Unrecognized code - passing through"
        End If
    End If
End Function

Function ConvertValue(rawValue As Variant, conversionRule As String, arCode As String) As Variant
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
    IsValidCode = True
    
    Select Case arCode
        Case "AR21", "AR189"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "5" Or code = "ND")
        Case "AR22", "AR23"
            IsValidCode = (code = "Y" Or code = "N" Or code = "ND")
        Case "AR84"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "ND")
        Case "AR105"
            IsValidCode = (code = "1" Or code = "2" Or code = "3" Or code = "4" Or code = "ND")
    End Select
End Function

Sub LogError(arCode As String, rowNum As Long, rawValue As Variant, errorMsg As String)
    On Error Resume Next
    Dim errorEntry As String
    errorEntry = "Row " & rowNum & " | " & arCode & " | Value: '" & rawValue & "' | " & errorMsg
    errorLog.Add errorEntry
End Sub

Sub ViewErrorLog()
    If errorLog Is Nothing Then
        MsgBox "No error log. Run mapping first.", vbInformation
        Exit Sub
    End If
    
    If errorLog.Count = 0 Then
        MsgBox "✓ No errors logged", vbInformation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim wsName As String
    wsName = "Data Quality Log"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(wsName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = wsName
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ws.Range("A1").Value = "Row"
    ws.Range("B1").Value = "AR Code"
    ws.Range("C1").Value = "Raw Value"
    ws.Range("D1").Value = "Issue"
    ws.Range("A1:D1").Font.Bold = True
    
    Dim i As Long
    For i = 1 To errorLog.Count
        Dim parts() As String
        parts = Split(errorLog(i), " | ")
        If UBound(parts) >= 3 Then
            ws.Cells(i + 1, 1).Value = Replace(parts(0), "Row ", "")
            ws.Cells(i + 1, 2).Value = parts(1)
            ws.Cells(i + 1, 3).Value = Replace(Replace(parts(2), "Value: '", ""), "'", "")
            ws.Cells(i + 1, 4).Value = parts(3)
        End If
    Next i
    
    ws.Columns("A:D").AutoFit
    ws.Activate
    
    MsgBox "Error log: " & errorLog.Count & " issues" & vbCrLf & _
           "See sheet '" & wsName & "'", vbInformation
End Sub

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
    ' Updated for v6.0 with new fields
    dict("AR3") = "Loan ID|MANDATORY"
    dict("AR5") = "Originator|Adj"
    dict("AR19") = "Number of Debtors|Joint Logic|NEW"
    dict("AR21") = "Employment|Adj"
    dict("AR22") = "FTB|Adj"
    dict("AR23") = "RTB|Adj"
    dict("AR26") = "Primary Income|Adj"
    dict("AR27") = "Income Verification|Adj"
    dict("AR28") = "Secondary Income|Adj"
    dict("AR31") = "CCJ Satisfied B1|Adj|NEW"
    dict("AR32") = "CCJ Value B1|Adj|NEW"
    dict("AR33") = "CCJ Unsatisfied B1|Adj"
    dict("AR35") = "Last CCJ Date B1|Adj"
    dict("AR36") = "Bankruptcy|Adj"
    dict("AR47") = "Repossession|Adj"
    dict("AR55") = "Origination Date|MANDATORY"
    dict("AR56") = "Maturity Date|Adj"
    dict("AR59") = "Purpose|Adj"
    dict("AR60") = "Shared Ownership|Adj|LTV"
    dict("AR66") = "Original Balance|MANDATORY"
    dict("AR67") = "Current Balance|MANDATORY"
    dict("AR69") = "Payment Method|Adj"
    dict("AR73") = "DTI|Adj"
    dict("AR84") = "Lien|Adj"
    dict("AR90") = "CCJ Value B2|Joint|NEW"
    dict("AR91") = "CCJ Satisfied B2|Joint|NEW"
    dict("AR92") = "Last CCJ Date B2|Joint|NEW"
    dict("AR103") = "Ownership %|LTV|NEW"
    dict("AR104") = "Borrower Property Value|LTV|NEW"
    dict("AR105") = "HTB|Adj"
    dict("AR107") = "Rate Type|Adj"
    dict("AR130") = "Occupancy|Adj"
    dict("AR131") = "Property Type|Validation"
    dict("AR136") = "Original Property Value|MANDATORY"
    dict("AR141") = "Current LTV|Adj"
    dict("AR143") = "Current Property Value|MANDATORY"
    dict("AR154") = "Rental Income|BTL"
    dict("AR156") = "DSCR|BTL"
    dict("AR166") = "Account Status|Status"
    dict("AR170") = "Months in Arrears|Adj"
    dict("AR173") = "Performance Arr|Adj"
End Sub

Function GenerateReport(mappedFields As Long, mappedLoans As Long, _
                       elapsedTime As Double, totalConfigured As Long, _
                       criticalMissing As Object) As String
    Dim report As String
    
    report = "✓ MAPPING COMPLETE!" & vbCrLf & vbCrLf
    report = report & "═══════════════════" & vbCrLf
    report = report & "STATISTICS" & vbCrLf
    report = report & "═══════════════════" & vbCrLf
    report = report & "Fields: " & mappedFields & " of " & totalConfigured & vbCrLf
    report = report & "Loans: " & Format(mappedLoans, "#,##0") & vbCrLf
    report = report & "Time: " & Format(elapsedTime, "0.0") & " sec" & vbCrLf
    report = report & "Rate: " & Format(mappedLoans / elapsedTime, "#,##0") & " loans/sec" & vbCrLf
    
    If criticalMissing.Count = 0 Then
        report = report & vbCrLf & "✓ ALL CRITICAL FIELDS PRESENT"
    Else
        report = report & vbCrLf & "⚠️ MISSING: " & criticalMissing.Count & " fields"
    End If
    
    GenerateReport = report
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
    MsgBox "S&P WAFF/WALS MODEL v6.0" & vbCrLf & vbCrLf & _
           "Release: 19-Jan-2026" & vbCrLf & vbCrLf & _
           "FEATURES:" & vbCrLf & _
           "✓ 52 fields sorted by AR code" & vbCrLf & _
           "✓ 8 new fields for CCJ/Shared Own" & vbCrLf & _
           "✓ Robust error handling" & vbCrLf & _
           "✓ CCJ joint calculation" & vbCrLf & _
           "✓ Shared ownership LTV" & vbCrLf & vbCrLf & _
           "Run 'ViewErrorLog' for data quality issues", _
           vbInformation, "Model v6.0"
End Sub

Sub ClearMappedData()
    Dim response As VbMsgBoxResult
    response = MsgBox("Clear ALL mapped data?", vbQuestion + vbYesNo)
    
    If response = vbYes Then
        Dim wsBoE As Worksheet
        Set wsBoE = ThisWorkbook.Sheets("Loan Tape (BoE)")
        wsBoE.Range("A5:AZ" & wsBoE.Cells(wsBoE.Rows.Count, 1).End(xlUp).Row).ClearContents
        MsgBox "✓ Data cleared", vbInformation
    End If
End Sub
