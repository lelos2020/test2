Attribute VB_Name = "BoE_Auto_Mapper"
Option Explicit

' ===================================================================
' BOE AUTO-MAPPER - CORRECTED BOE CODE CONVERSIONS
' Version 5.2.2 - Based on ACTUAL BoE PRA110 Template Dictionary
' ===================================================================
'
' CHANGELOG v5.2.2:
' - CORRECTED all BoE code conversions based on official template
' - CRITICAL: AR60 Shared Ownership: 1 = NO (not shared), 2-6 = YES
' - CRITICAL: AR105 Help to Buy: 1-3 = YES, 4 = NO  
' - CRITICAL: AR173 Performance Arrangement: Date field check
' - Maps ALL 233 AR fields from BoE template
' - Comprehensive 38 critical field tracking
'
' ===================================================================

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
    
    ' Remove existing buttons
    ws.Buttons.Delete
    
    ' Add new button
    Set btn = ws.Buttons.Add(10, 10, 200, 40)
    btn.OnAction = "MapBoEFields"
    btn.Text = "Map BoE Fields" & vbCrLf & "(AR Codes)"
    btn.Font.Bold = True
    btn.Font.Size = 11
    btn.Font.Color = RGB(255, 255, 255)
    
    ' Style the button
    With btn
        .ShapeRange.Fill.ForeColor.RGB = RGB(54, 96, 146)
        .ShapeRange.Line.ForeColor.RGB = RGB(54, 96, 146)
    End With
    
    MsgBox "✓ Button Added Successfully!" & vbCrLf & vbCrLf & _
           "Model Version: 5.2.2 - CORRECTED BoE Conversions" & vbCrLf & vbCrLf & _
           "CRITICAL FIXES:" & vbCrLf & _
           "• AR60: 1=No (NOT shared), 2-6=Yes (IS shared)" & vbCrLf & _
           "• AR105: 1-3=Yes, 4=No" & vbCrLf & _
           "• All 233 AR fields supported" & vbCrLf & vbCrLf & _
           "Paste BoE PRA110 loan tape and click 'Map BoE Fields'", _
           vbInformation, "BoE Auto-Mapper v5.2.2"
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
    
    Set wsRaw = ThisWorkbook.Sheets("Loan Tape (BoE Raw)")
    Set wsMapper = ThisWorkbook.Sheets("BoE Auto-Mapper")
    Set wsBoE = ThisWorkbook.Sheets("Loan Tape (BoE)")
    
    ' ================================================================
    ' STEP 1: FIND HEADER ROW WITH AR CODES
    ' ================================================================
    headerRow = FindARCodeRow(wsRaw)
    
    If headerRow = 0 Then
        MsgBox "❌ AR Codes Not Found" & vbCrLf & vbCrLf & _
               "Cannot find Bank of England AR field codes in first 20 rows." & vbCrLf & vbCrLf & _
               "Required format: AR1, AR3, AR55, etc." & vbCrLf & vbCrLf & _
               "Per BoE PRA110 template: AR codes should be in Row 1", _
               vbCritical, "Header Validation Failed"
        GoTo CleanUp
    End If
    
    ' ================================================================
    ' STEP 2: BUILD MAPPING DICTIONARY
    ' ================================================================
    Set colMapping = BuildMappingDictionary(wsMapper)
    
    If colMapping.Count = 0 Then
        MsgBox "❌ No Field Mappings" & vbCrLf & vbCrLf & _
               "BoE Auto-Mapper sheet configuration is empty or invalid.", _
               vbCritical, "Configuration Error"
        GoTo CleanUp
    End If
    
    ' ================================================================
    ' STEP 3: VALIDATE DATA DIMENSIONS
    ' ================================================================
    lastCol = wsRaw.Cells(headerRow, wsRaw.Columns.Count).End(xlToLeft).Column
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= headerRow Then
        MsgBox "❌ No Loan Data Found" & vbCrLf & vbCrLf & _
               "Please paste your loan tape below the AR code headers.", _
               vbExclamation, "No Data"
        GoTo CleanUp
    End If
    
    mappedLoans = lastRow - headerRow
    
    ' Large file notification
    If mappedLoans > 10000 Then
        MsgBox "Processing " & Format(mappedLoans, "#,##0") & " loans..." & vbCrLf & _
               "Please wait (approximately " & Format(mappedLoans / 500, "0") & " seconds)", _
               vbInformation, "Large File Processing"
    End If
    
    ' ================================================================
    ' STEP 4: CLEAR TARGET SHEET
    ' ================================================================
    wsBoE.Range("A5:AZ" & wsBoE.Cells(wsBoE.Rows.Count, 1).End(xlUp).Row).ClearContents
    
    ' ================================================================
    ' STEP 5: MAP FIELDS WITH CRITICAL FIELD TRACKING
    ' ================================================================
    Set criticalMissing = CreateObject("Scripting.Dictionary")
    InitializeCriticalFields criticalMissing
    
    mappedFields = 0
    
    ' Process each column
    For j = 1 To lastCol
        Dim headerText As String, extractedCode As String
        headerText = UCase(Trim(wsRaw.Cells(headerRow, j).Value))
        extractedCode = ExtractARCode(headerText)
        
        If extractedCode <> "" And colMapping.Exists(extractedCode) Then
            boEColumn = colMapping(extractedCode)(0)
            conversionRule = colMapping(extractedCode)(1)
            
            ' Copy and convert data
            For i = headerRow + 1 To lastRow
                Dim rawValue As Variant, convertedValue As Variant
                rawValue = wsRaw.Cells(i, j).Value
                convertedValue = ConvertValue(rawValue, conversionRule, extractedCode)
                wsBoE.Cells(i - headerRow + 4, boEColumn).Value = convertedValue
            Next i
            
            mappedFields = mappedFields + 1
            
            ' Remove from critical missing list if found
            If criticalMissing.Exists(extractedCode) Then
                criticalMissing.Remove extractedCode
            End If
        End If
    Next j
    
    ' ================================================================
    ' STEP 6: GENERATE COMPREHENSIVE REPORT
    ' ================================================================
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    Dim reportMsg As String
    reportMsg = GenerateDataQualityReport(mappedFields, mappedLoans, _
                                         elapsedTime, colMapping.Count, _
                                         criticalMissing)
    
    ' Show report with appropriate severity
    If criticalMissing.Count = 0 Then
        MsgBox reportMsg, vbInformation, "Mapping Complete - Excellent Data Quality"
    ElseIf criticalMissing.Count <= 5 Then
        MsgBox reportMsg, vbExclamation, "Mapping Complete - Minor Issues"
    Else
        MsgBox reportMsg, vbCritical, "Mapping Complete - Critical Issues Detected"
    End If
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ ERROR: " & Err.Description & vbCrLf & _
           "Error #" & Err.Number & vbCrLf & vbCrLf & _
           "Please check loan tape format and model structure.", _
           vbCritical, "Processing Error"
    GoTo CleanUp
End Sub

' ===================================================================
' HELPER FUNCTIONS
' ===================================================================

Function FindARCodeRow(ws As Worksheet) As Long
    ' Find row containing AR codes in first 20 rows
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
    ' Build dictionary of AR code → (column, conversion rule) mappings
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
    ' Initialize dictionary with ALL critical fields for FF calculations
    
    ' MANDATORY BASE (6)
    dict("AR3") = "Loan Identifier|MANDATORY"
    dict("AR66") = "Original Balance|MANDATORY"
    dict("AR67") = "Current Balance|MANDATORY"
    dict("AR136") = "Original Property Value|MANDATORY"
    dict("AR143") = "Current Property Value|MANDATORY"
    dict("AR55") = "Origination Date|MANDATORY"
    
    ' ADJUSTMENT DRIVERS (8)
    dict("AR5") = "Originator|Adj: Originator"
    dict("AR131") = "Property Type|Validation"
    dict("AR130") = "Occupancy Type|Adj: BTL/1.7x if missing"
    dict("AR141") = "Current LTV|Adj: LTV Curve"
    dict("AR84") = "Lien Position|Adj: 1.50x Second Lien"
    dict("AR59") = "Loan Purpose|Adj: Purpose"
    dict("AR69") = "Payment Method|Adj: Interest Only"
    dict("AR56") = "Maturity Date|Calc: Remaining Term"
    
    ' BORROWER (4)
    dict("AR26") = "Primary Income|Adj: Income Multiple"
    dict("AR28") = "Secondary Income|Adj: Income Multiple"
    dict("AR73") = "Debt to Income|Alt: Income Multiple"
    dict("AR19") = "Number of Debtors|Logic: Joint Thresholds"
    
    ' PRODUCT FEATURES (4)
    dict("AR22") = "First Time Buyer|Adj: FTB"
    dict("AR105") = "Help to Buy|Adj: 1.35x - CRITICAL"
    dict("AR60") = "Shared Ownership|Adj: Shared Own"
    dict("AR23") = "Right to Buy|Adj: RTB"
    
    ' CREDIT HISTORY (5)
    dict("AR31") = "CCJ Satisfied|Adj: CCJ Tiers"
    dict("AR33") = "CCJ Unsatisfied|Adj: CCJ Tiers"
    dict("AR35") = "Last CCJ Date|Adj: CCJ Age Tiers"
    dict("AR36") = "Bankruptcy/IVA|Adj: 1.50x-3.00x"
    dict("AR47") = "Prior Repossession|Adj: Repossession"
    
    ' PERFORMANCE (6)
    dict("AR166") = "Account Status|Status Check"
    dict("AR169") = "Arrears Balance|Arrears Amount"
    dict("AR170") = "Months in Arrears|Adj: 1.00x-15.00x"
    dict("AR173") = "Performance Arrangement|Forbearance Flag"
    dict("AR167") = "Date Last Current|Re-performing Check"
    dict("AR168") = "Date Last in Arrears|Re-performing Check"
    
    ' BTL SPECIFIC (2)
    dict("AR154") = "Rental Income|Calc: DSCR"
    dict("AR156") = "DSCR|Adj: BTL/DSCR"
    
    ' EMPLOYMENT (2)
    dict("AR21") = "Employment Status|Adj: Self-Employed"
    dict("AR27") = "Income Verification|Adj: Self-Cert"
    
    ' PAYMENT SHOCK (1)
    dict("AR107") = "Interest Rate Type|Adj: Payment Shock"
End Function

Function GenerateDataQualityReport(mappedFields As Long, mappedLoans As Long, _
                                   elapsedTime As Double, totalConfigured As Long, _
                                   criticalMissing As Object) As String
    ' Generate comprehensive data quality report
    
    Dim report As String
    Dim coveragePct As Double
    
    report = "✓ MAPPING COMPLETE!" & vbCrLf & vbCrLf
    report = report & "═══════════════════════════════" & vbCrLf
    report = report & "MAPPING STATISTICS" & vbCrLf
    report = report & "═══════════════════════════════" & vbCrLf
    report = report & "Fields Mapped: " & mappedFields & " of " & totalConfigured & " configured" & vbCrLf
    report = report & "Loans Processed: " & Format(mappedLoans, "#,##0") & vbCrLf
    report = report & "Processing Time: " & Format(elapsedTime, "0.0") & " sec" & vbCrLf
    report = report & "Processing Rate: " & Format(mappedLoans / elapsedTime, "#,##0") & " loans/sec" & vbCrLf
    
    coveragePct = (mappedFields / totalConfigured) * 100
    report = report & "Field Coverage: " & Format(coveragePct, "0.0") & "%" & vbCrLf
    
    ' Critical fields analysis
    report = report & vbCrLf & "═══════════════════════════════" & vbCrLf
    report = report & "DATA QUALITY ASSESSMENT" & vbCrLf
    report = report & "═══════════════════════════════" & vbCrLf
    
    If criticalMissing.Count = 0 Then
        report = report & "✓ ALL CRITICAL FIELDS PRESENT" & vbCrLf
        report = report & "✓ Model ready for FF calculations" & vbCrLf
    Else
        report = report & "⚠️ MISSING CRITICAL FIELDS: " & criticalMissing.Count & vbCrLf & vbCrLf
        
        ' Categorize missing fields
        Dim mandatoryMissing As String, adjustmentMissing As String, otherMissing As String
        mandatoryMissing = ""
        adjustmentMissing = ""
        otherMissing = ""
        
        Dim key As Variant
        For Each key In criticalMissing.Keys()
            Dim fieldInfo As String
            fieldInfo = criticalMissing(key)
            
            If InStr(fieldInfo, "MANDATORY") > 0 Then
                mandatoryMissing = mandatoryMissing & "  " & key & " - " & Split(fieldInfo, "|")(0) & vbCrLf
            ElseIf InStr(fieldInfo, "Adj:") > 0 Then
                adjustmentMissing = adjustmentMissing & "  " & key & " - " & Split(fieldInfo, "|")(0) & vbCrLf
            Else
                otherMissing = otherMissing & "  " & key & " - " & Split(fieldInfo, "|")(0) & vbCrLf
            End If
        Next key
        
        If mandatoryMissing <> "" Then
            report = report & "MANDATORY (Cannot calculate):" & vbCrLf & mandatoryMissing & vbCrLf
        End If
        
        If adjustmentMissing <> "" Then
            report = report & "Adjustment Fields (Default 1.0x):" & vbCrLf & adjustmentMissing & vbCrLf
        End If
        
        If otherMissing <> "" Then
            report = report & "Supporting Fields:" & vbCrLf & otherMissing & vbCrLf
        End If
        
        report = report & "⚠️ Model may calculate but with reduced accuracy" & vbCrLf
    End If
    
    ' Next steps
    report = report & vbCrLf & "NEXT STEPS:" & vbCrLf
    report = report & "1. Review 'Loan Tape (BoE)' mapped data" & vbCrLf
    report = report & "2. Check 'Pool Summary' calculations" & vbCrLf
    report = report & "3. Review 'Adjustment Overview' dashboard" & vbCrLf
    
    If criticalMissing.Count > 0 Then
        report = report & "4. ⚠️ Address missing critical fields if possible" & vbCrLf
    End If
    
    GenerateDataQualityReport = report
End Function

Function ExtractARCode(headerText As String) As String
    ' Extract AR code from various header formats
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

Function ConvertValue(rawValue As Variant, conversionRule As String, arCode As String) As Variant
    ' Convert raw BoE codes to S&P format values
    ' v5.2.2: CORRECTED CONVERSIONS BASED ON BOE PRA110 TEMPLATE
    
    If IsEmpty(rawValue) Or rawValue = "" Or IsNull(rawValue) Then
        ConvertValue = ""
        Exit Function
    End If
    
    If conversionRule = "Direct" Or conversionRule = "" Then
        ConvertValue = rawValue
        Exit Function
    End If
    
    Dim code As String
    code = CStr(rawValue)
    
    Select Case True
        ' ================================================================
        ' AR22: FIRST TIME BUYER - Y/N/ND format
        ' ================================================================
        Case arCode = "AR22"
            Select Case UCase(code)
                Case "Y": ConvertValue = "Yes"
                Case "N": ConvertValue = "No"
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = "No"
            End Select
            
        ' ================================================================
        ' AR23: RIGHT TO BUY - Y/N/ND format
        ' ================================================================
        Case arCode = "AR23"
            Select Case UCase(code)
                Case "Y": ConvertValue = "Yes"
                Case "N": ConvertValue = "No"
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = "No"
            End Select
            
        ' ================================================================
        ' AR60: SHARED OWNERSHIP - CRITICAL FIX!
        ' BoE: 1=NOT shared, 2-6=IS shared
        ' ================================================================
        Case arCode = "AR60"
            Select Case code
                Case "1": ConvertValue = "No"   ' NOT Shared Ownership
                Case "2": ConvertValue = "Yes"  ' Central Government
                Case "3": ConvertValue = "Yes"  ' Local Government
                Case "4": ConvertValue = "Yes"  ' Housing Associations
                Case "5": ConvertValue = "Yes"  ' Building Developers
                Case "6": ConvertValue = "Yes"  ' Other
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = "No"
            End Select
            
        ' ================================================================
        ' AR69: REPAYMENT METHOD
        ' ================================================================
        Case arCode = "AR69"
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
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        ' ================================================================
        ' AR84: LIEN POSITION
        ' ================================================================
        Case arCode = "AR84"
            Select Case code
                Case "1": ConvertValue = "First"
                Case "2": ConvertValue = "Second"
                Case "3": ConvertValue = "Third"
                Case "4": ConvertValue = "Other"
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        ' ================================================================
        ' AR105: HELP TO BUY - CRITICAL FIX!
        ' BoE: 1-3=Yes, 4=No
        ' ================================================================
        Case arCode = "AR105"
            Select Case code
                Case "1": ConvertValue = "Yes"  ' Help to Buy 1 (Equity)
                Case "2": ConvertValue = "Yes"  ' Help to Buy 2 (MIG)
                Case "3": ConvertValue = "Yes"  ' Other
                Case "4": ConvertValue = "No"   ' No or not applicable
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = "No"
            End Select
            
        ' ================================================================
        ' AR107: INTEREST RATE TYPE
        ' ================================================================
        Case arCode = "AR107"
            Select Case code
                Case "1": ConvertValue = "Variable"  ' Floating for life
                Case "2": ConvertValue = "Tracker"   ' Linked to index
                Case "3": ConvertValue = "Fixed"     ' Fixed for life
                Case "4": ConvertValue = "Fixed"     ' Fixed with resets
                Case "5": ConvertValue = "Fixed"     ' Fixed with switch
                Case "6": ConvertValue = "Variable"  ' Capped
                Case "7": ConvertValue = "Discount"
                Case "8": ConvertValue = "Other"
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        ' ================================================================
        ' AR130: OCCUPANCY TYPE - CRITICAL: Blank = 1.7x
        ' ================================================================
        Case arCode = "AR130"
            Select Case code
                Case "1": ConvertValue = "Owner Occupied"
                Case "2": ConvertValue = "Partially Owner Occupied"
                Case "3": ConvertValue = "Buy to Let"
                Case "4": ConvertValue = "Second Home"
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = ""  ' Blank triggers 1.7x
            End Select
            
        ' ================================================================
        ' AR131: PROPERTY TYPE
        ' ================================================================
        Case arCode = "AR131"
            Select Case code
                Case "1": ConvertValue = "House"
                Case "2": ConvertValue = "Flat"
                Case "3": ConvertValue = "Bungalow"
                Case "4": ConvertValue = "Terraced House"
                Case "5": ConvertValue = "Multifamily (with recourse)"
                Case "6": ConvertValue = "Multifamily (without recourse)"
                Case "7": ConvertValue = "Partially Commercial"
                Case "ND": ConvertValue = "No Data"
                Case Else: ConvertValue = ""
            End Select
            
        ' ================================================================
        ' AR173: PERFORMANCE ARRANGEMENT - CRITICAL FIX!
        ' Date field - if date exists = Yes, otherwise No
        ' ================================================================
        Case arCode = "AR173"
            If IsDate(rawValue) Then
                ConvertValue = "Yes"
            ElseIf UCase(code) = "ND" Then
                ConvertValue = "No Data"
            Else
                ConvertValue = "No"
            End If
            
        ' ================================================================
        ' NUMERIC/PERCENTAGE
        ' ================================================================
        Case InStr(conversionRule, "Percentage") > 0 Or InStr(conversionRule, "Number") > 0
            If IsNumeric(rawValue) Then
                ConvertValue = CDbl(rawValue)
            Else
                ConvertValue = ""
            End If
            
        ' ================================================================
        ' DATE
        ' ================================================================
        Case InStr(conversionRule, "Date") > 0
            If IsDate(rawValue) Then
                ConvertValue = CDate(rawValue)
            Else
                ConvertValue = ""
            End If
            
        ' ================================================================
        ' DEFAULT - Pass through
        ' ================================================================
        Case Else
            ConvertValue = rawValue
    End Select
End Function

Function ColumnLetterToNumber(columnLetter As String) As Long
    On Error Resume Next
    ColumnLetterToNumber = Range(columnLetter & "1").Column
    If Err.Number <> 0 Then ColumnLetterToNumber = 0
    On Error GoTo 0
End Function

' ===================================================================
' UTILITY SUBS
' ===================================================================

Sub TestMapping()
    Dim wsMapper As Worksheet
    Set wsMapper = ThisWorkbook.Sheets("BoE Auto-Mapper")
    
    Dim mappingCount As Long, criticalCount As Long
    Dim i As Long, arCode As String
    
    For i = 5 To wsMapper.Cells(wsMapper.Rows.Count, 1).End(xlUp).Row
        arCode = wsMapper.Cells(i, 1).Value
        If arCode <> "" And Left(arCode, 2) = "AR" Then
            mappingCount = mappingCount + 1
            If IsCriticalField(arCode) Then criticalCount = criticalCount + 1
        End If
    Next i
    
    MsgBox "✓ Configuration Test Complete" & vbCrLf & vbCrLf & _
           "Total Mappings: " & mappingCount & vbCrLf & _
           "Critical Fields: " & criticalCount & " of 38" & vbCrLf & vbCrLf & _
           "Model: v5.2.2 - CORRECTED BoE Conversions" & vbCrLf & vbCrLf & _
           "CRITICAL FIXES:" & vbCrLf & _
           "• AR60: 1=No, 2-6=Yes" & vbCrLf & _
           "• AR105: 1-3=Yes, 4=No" & vbCrLf & _
           "• AR173: Date check" & vbCrLf & _
           "• All 233 AR fields supported", _
           vbInformation, "Configuration Test"
End Sub

Function IsCriticalField(arCode As String) As Boolean
    Dim criticalFields As Variant
    criticalFields = Array("AR3", "AR5", "AR19", "AR21", "AR22", "AR23", "AR26", "AR27", "AR28", _
                          "AR31", "AR33", "AR35", "AR36", "AR47", "AR55", "AR56", "AR59", "AR60", _
                          "AR66", "AR67", "AR69", "AR73", "AR84", "AR105", "AR107", "AR130", "AR131", _
                          "AR136", "AR141", "AR143", "AR154", "AR156", "AR166", "AR167", "AR168", _
                          "AR169", "AR170", "AR173")
    
    Dim i As Long
    For i = LBound(criticalFields) To UBound(criticalFields)
        If criticalFields(i) = arCode Then
            IsCriticalField = True
            Exit Function
        End If
    Next i
    IsCriticalField = False
End Function

Sub ShowModelInfo()
    MsgBox "S&P WAFF/WALS MODEL" & vbCrLf & vbCrLf & _
           "Version: 5.2.2" & vbCrLf & _
           "Status: CORRECTED BoE Conversions" & vbCrLf & _
           "Release: 19-Jan-2026" & vbCrLf & vbCrLf & _
           "CRITICAL CORRECTIONS:" & vbCrLf & _
           "✓ AR60 Shared Own: 1=No, 2-6=Yes" & vbCrLf & _
           "✓ AR105 Help to Buy: 1-3=Yes, 4=No" & vbCrLf & _
           "✓ AR173 Perf Arrange: Date check" & vbCrLf & _
           "✓ All 233 AR fields mapped" & vbCrLf & vbCrLf & _
           "Based on BoE PRA110 Template Dictionary", _
           vbInformation, "Model Information"
End Sub

Sub ClearMappedData()
    Dim response As VbMsgBoxResult
    response = MsgBox("Clear ALL mapped data from 'Loan Tape (BoE)'?", _
                     vbQuestion + vbYesNo, "Clear Data")
    
    If response = vbYes Then
        Dim wsBoE As Worksheet
        Set wsBoE = ThisWorkbook.Sheets("Loan Tape (BoE)")
        wsBoE.Range("A5:AZ" & wsBoE.Cells(wsBoE.Rows.Count, 1).End(xlUp).Row).ClearContents
        MsgBox "✓ Data cleared", vbInformation
    End If
End Sub
