Attribute VB_Name = "BoE_Auto_Mapper"
Option Explicit

' ===================================================================
' BOE AUTO-MAPPER - AUTOMATIC FIELD MAPPING FOR S&P WAFF/WALS MODEL
' Version 5.2 - Updated for S&P Methodology Compliance
' ===================================================================
'
' CHANGELOG v5.2:
' - Updated code conversion logic for methodology compliance
' - Enhanced error handling and validation
' - Added mapping statistics and data quality checks
' - Optimized for large loan tapes (50K+ loans)
' - Compatible with AR codes from BoE PRA110 template
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
        .ShapeRange.Fill.ForeColor.RGB = RGB(54, 96, 146) ' Blue
        .ShapeRange.Line.ForeColor.RGB = RGB(54, 96, 146)
    End With
    
    MsgBox "✓ Button Added Successfully!" & vbCrLf & vbCrLf & _
           "Instructions:" & vbCrLf & _
           "1. Paste your loan tape into 'Loan Tape (BoE Raw)'" & vbCrLf & _
           "2. Ensure AR codes are in the header row (AR3, AR55, AR131, etc.)" & vbCrLf & _
           "3. Click the 'Map BoE Fields' button" & vbCrLf & _
           "4. Review mapping results in 'Loan Tape (BoE)'" & vbCrLf & vbCrLf & _
           "Model Version: 5.2 - S&P Methodology Compliant", _
           vbInformation, "BoE Auto-Mapper v5.2"
End Sub

Sub MapBoEFields()
    Dim wsRaw As Worksheet, wsMapper As Worksheet, wsBoE As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim arCode As String, boEColumn As Long, conversionRule As String
    Dim colMapping As Object, headerRow As Long
    Dim startTime As Double
    Dim mappedFields As Long, mappedLoans As Long
    Dim missingFields As String, dataQualityWarnings As String
    
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' Set worksheet references
    Set wsRaw = ThisWorkbook.Sheets("Loan Tape (BoE Raw)")
    Set wsMapper = ThisWorkbook.Sheets("BoE Auto-Mapper")
    Set wsBoE = ThisWorkbook.Sheets("Loan Tape (BoE)")
    
    ' ================================================================
    ' STEP 1: FIND HEADER ROW WITH AR CODES
    ' ================================================================
    headerRow = 0
    For i = 1 To 20
        If InStr(1, wsRaw.Cells(i, 1).Value, "AR", vbTextCompare) > 0 Then
            headerRow = i
            Exit For
        End If
    Next i
    
    If headerRow = 0 Then
        MsgBox "❌ AR Codes Not Found" & vbCrLf & vbCrLf & _
               "Cannot find Bank of England field codes (AR3, AR55, AR131, etc.) " & _
               "in the first 20 rows of your loan tape." & vbCrLf & vbCrLf & _
               "Please ensure:" & vbCrLf & _
               "• Your loan tape has AR codes in the header row" & vbCrLf & _
               "• AR codes follow BoE PRA110 format (e.g., AR3, AR131, AR166)" & vbCrLf & _
               "• Headers are in the first 20 rows of the sheet", _
               vbCritical, "Header Row Not Found"
        GoTo CleanUp
    End If
    
    ' ================================================================
    ' STEP 2: BUILD MAPPING DICTIONARY FROM BOE AUTO-MAPPER
    ' ================================================================
    Set colMapping = CreateObject("Scripting.Dictionary")
    
    Dim mapperLastRow As Long
    mapperLastRow = wsMapper.Cells(wsMapper.Rows.Count, 1).End(xlUp).Row
    
    ' Read mapping configuration (starting from row 5 after headers)
    For i = 5 To mapperLastRow
        arCode = Trim(wsMapper.Cells(i, 1).Value) ' Column A: AR Code
        If arCode <> "" And Left(arCode, 2) = "AR" Then
            boEColumn = ColumnLetterToNumber(wsMapper.Cells(i, 3).Value) ' Column C: Target column
            conversionRule = wsMapper.Cells(i, 6).Value ' Column F: Conversion rule
            
            If boEColumn > 0 Then
                colMapping(UCase(arCode)) = Array(boEColumn, conversionRule)
            End If
        End If
    Next i
    
    If colMapping.Count = 0 Then
        MsgBox "❌ No Field Mappings Configured" & vbCrLf & vbCrLf & _
               "The BoE Auto-Mapper sheet has no valid field mappings." & vbCrLf & vbCrLf & _
               "Please check that:" & vbCrLf & _
               "• BoE Auto-Mapper sheet contains AR code mappings" & vbCrLf & _
               "• Column A has AR codes (AR3, AR55, etc.)" & vbCrLf & _
               "• Column C has target column letters" & vbCrLf & _
               "• Column F has conversion rules", _
               vbCritical, "Configuration Error"
        GoTo CleanUp
    End If
    
    ' ================================================================
    ' STEP 3: FIND DATA DIMENSIONS
    ' ================================================================
    lastCol = wsRaw.Cells(headerRow, wsRaw.Columns.Count).End(xlToLeft).Column
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= headerRow Then
        MsgBox "❌ No Data Found" & vbCrLf & vbCrLf & _
               "No loan data found below the header row." & vbCrLf & vbCrLf & _
               "Please paste your loan tape data below the AR code headers.", _
               vbExclamation, "No Data"
        GoTo CleanUp
    End If
    
    mappedLoans = lastRow - headerRow
    
    ' Progress notification for large files
    If mappedLoans > 10000 Then
        MsgBox "Processing " & Format(mappedLoans, "#,##0") & " loans..." & vbCrLf & _
               "This may take a minute. Please wait.", vbInformation, "Large File Detected"
    End If
    
    ' ================================================================
    ' STEP 4: CLEAR EXISTING DATA IN TARGET SHEET
    ' ================================================================
    ' Clear from row 5 onwards (preserve headers in rows 1-4)
    wsBoE.Range("A5:AZ" & wsBoE.Cells(wsBoE.Rows.Count, 1).End(xlUp).Row).ClearContents
    
    ' ================================================================
    ' STEP 5: MAP FIELDS AND CONVERT DATA
    ' ================================================================
    mappedFields = 0
    missingFields = ""
    dataQualityWarnings = ""
    
    ' Track critical missing fields
    Dim criticalFields As Object
    Set criticalFields = CreateObject("Scripting.Dictionary")
    criticalFields("AR3") = "Loan Identifier"
    criticalFields("AR66") = "Original Balance"
    criticalFields("AR67") = "Current Balance"
    criticalFields("AR131") = "Property Type"
    criticalFields("AR130") = "Occupancy Type"
    criticalFields("AR141") = "Current LTV"
    
    ' Process each column in the raw tape
    For j = 1 To lastCol
        Dim headerText As String
        headerText = UCase(Trim(wsRaw.Cells(headerRow, j).Value))
        
        ' Extract AR code from header
        Dim extractedCode As String
        extractedCode = ExtractARCode(headerText)
        
        If extractedCode <> "" And colMapping.Exists(extractedCode) Then
            ' Get target column and conversion rule
            boEColumn = colMapping(extractedCode)(0)
            conversionRule = colMapping(extractedCode)(1)
            
            ' Copy and convert data row by row
            For i = headerRow + 1 To lastRow
                Dim rawValue As Variant
                Dim convertedValue As Variant
                
                rawValue = wsRaw.Cells(i, j).Value
                convertedValue = ConvertValue(rawValue, conversionRule, extractedCode)
                
                ' Write to target column (offset by 4 to start at row 5)
                wsBoE.Cells(i - headerRow + 4, boEColumn).Value = convertedValue
            Next i
            
            mappedFields = mappedFields + 1
            
            ' Remove from critical fields list if found
            If criticalFields.Exists(extractedCode) Then
                criticalFields.Remove extractedCode
            End If
        End If
    Next j
    
    ' ================================================================
    ' STEP 6: DATA QUALITY CHECKS
    ' ================================================================
    
    ' Check for critical missing fields
    If criticalFields.Count > 0 Then
        Dim key As Variant
        dataQualityWarnings = "⚠️ CRITICAL FIELDS MISSING:" & vbCrLf
        For Each key In criticalFields.Keys()
            dataQualityWarnings = dataQualityWarnings & _
                "  • " & key & " (" & criticalFields(key) & ")" & vbCrLf
        Next key
        dataQualityWarnings = dataQualityWarnings & vbCrLf
    End If
    
    ' ================================================================
    ' STEP 7: CALCULATE ELAPSED TIME AND SHOW RESULTS
    ' ================================================================
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    Dim resultMsg As String
    resultMsg = "✓ MAPPING COMPLETE!" & vbCrLf & vbCrLf
    
    ' Mapping statistics
    resultMsg = resultMsg & "═══════════════════════════════" & vbCrLf
    resultMsg = resultMsg & "MAPPING STATISTICS" & vbCrLf
    resultMsg = resultMsg & "═══════════════════════════════" & vbCrLf
    resultMsg = resultMsg & "Fields Mapped: " & mappedFields & " of " & colMapping.Count & " configured" & vbCrLf
    resultMsg = resultMsg & "Loans Processed: " & Format(mappedLoans, "#,##0") & vbCrLf
    resultMsg = resultMsg & "Processing Time: " & Format(elapsedTime, "0.0") & " seconds" & vbCrLf
    resultMsg = resultMsg & "Processing Rate: " & Format(mappedLoans / elapsedTime, "#,##0") & " loans/sec" & vbCrLf
    
    ' Coverage percentage
    Dim coveragePct As Double
    coveragePct = (mappedFields / colMapping.Count) * 100
    resultMsg = resultMsg & "Field Coverage: " & Format(coveragePct, "0.0") & "%" & vbCrLf
    
    ' Add data quality warnings if any
    If dataQualityWarnings <> "" Then
        resultMsg = resultMsg & vbCrLf & dataQualityWarnings
    End If
    
    ' Next steps
    resultMsg = resultMsg & vbCrLf & "NEXT STEPS:" & vbCrLf
    resultMsg = resultMsg & "1. Review mapped data in 'Loan Tape (BoE)'" & vbCrLf
    resultMsg = resultMsg & "2. Check 'Pool Summary' for calculated WAFF" & vbCrLf
    resultMsg = resultMsg & "3. Review 'Adjustment Overview' for factor analysis" & vbCrLf
    
    ' Show appropriate message type
    If criticalFields.Count > 0 Then
        MsgBox resultMsg, vbExclamation, "Mapping Complete - Warnings Detected"
    Else
        MsgBox resultMsg, vbInformation, "BoE Auto-Mapper v5.2 - Success"
    End If
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "❌ ERROR OCCURRED" & vbCrLf & vbCrLf & _
           "Error Description: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Please check:" & vbCrLf & _
           "• Loan tape format is correct" & vbCrLf & _
           "• AR codes are properly formatted" & vbCrLf & _
           "• No merged cells in data area" & vbCrLf & _
           "• Sufficient memory for large files", _
           vbCritical, "Mapping Error"
    GoTo CleanUp
End Sub

Function ExtractARCode(headerText As String) As String
    ' Extract AR code from header text (e.g., "AR3", "AR131")
    ' Handles various formats: "AR3", "AR3 - Loan ID", "Loan ID (AR3)", etc.
    
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
    ' Convert raw value based on conversion rule and AR code
    ' Updated for v5.2 S&P methodology compliance
    
    ' Handle empty values
    If IsEmpty(rawValue) Or rawValue = "" Or IsNull(rawValue) Then
        ConvertValue = ""
        Exit Function
    End If
    
    ' Direct mapping - no conversion needed
    If conversionRule = "Direct" Or conversionRule = "" Then
        ConvertValue = rawValue
        Exit Function
    End If
    
    Dim code As String
    code = CStr(rawValue)
    
    ' Apply conversions based on the rule or AR code
    Select Case True
        ' ================================================================
        ' PROPERTY TYPE (AR131) - S&P Methodology Compliant
        ' ================================================================
        Case InStr(conversionRule, "House") > 0 Or arCode = "AR131"
            Select Case code
                Case "1": ConvertValue = "House"
                Case "2": ConvertValue = "Flat"
                Case "3": ConvertValue = "Bungalow"
                Case "4": ConvertValue = "Maisonette"
                Case Else: ConvertValue = ""  ' Blank for invalid codes
            End Select
            
        ' ================================================================
        ' OCCUPANCY TYPE (AR130) - S&P Methodology Compliant
        ' ================================================================
        Case InStr(conversionRule, "Owner Occupied") > 0 Or arCode = "AR130"
            Select Case code
                Case "1": ConvertValue = "Owner Occupied"
                Case "2": ConvertValue = "Buy to Let"
                Case "3": ConvertValue = "Second Home"
                Case "4": ConvertValue = "Investment"
                Case Else: ConvertValue = ""  ' Blank triggers 1.7x adjustment per v5.2
            End Select
            
        ' ================================================================
        ' INTEREST RATE TYPE (AR107) - S&P Methodology Compliant
        ' ================================================================
        Case InStr(conversionRule, "Fixed") > 0 Or arCode = "AR107"
            Select Case code
                Case "1": ConvertValue = "Fixed"
                Case "2": ConvertValue = "Variable"
                Case "3": ConvertValue = "Tracker"
                Case Else: ConvertValue = ""
            End Select
            
        ' ================================================================
        ' REPAYMENT METHOD (AR69) - S&P Methodology Compliant
        ' ================================================================
        Case InStr(conversionRule, "Repayment") > 0 Or arCode = "AR69"
            Select Case code
                Case "1": ConvertValue = "Repayment"
                Case "2": ConvertValue = "Interest Only"
                Case "3": ConvertValue = "Part and Part"
                Case Else: ConvertValue = ""
            End Select
            
        ' ================================================================
        ' LIEN POSITION (AR84) - S&P Methodology Compliant
        ' ================================================================
        Case arCode = "AR84"
            Select Case code
                Case "1": ConvertValue = "First"
                Case "2": ConvertValue = "Second"  ' 1.50x adjustment per v5.2
                Case "3": ConvertValue = "Third"
                Case Else: ConvertValue = ""
            End Select
            
        ' ================================================================
        ' YES/NO CONVERSIONS - Multiple patterns supported
        ' ================================================================
        
        ' Pattern 1: 1,2=Yes (Help to Buy AR105, Shared Ownership AR60, etc.)
        Case InStr(conversionRule, "1,2=Yes") > 0 Or _
             arCode = "AR105" Or arCode = "AR60" Or arCode = "AR173"
            If code = "1" Or code = "2" Then
                ConvertValue = "Yes"
            Else
                ConvertValue = "No"
            End If
            
        ' Pattern 2: 1=Yes (Right to Buy AR23, First Time Buyer AR22)
        Case InStr(conversionRule, "1=Yes") > 0 Or _
             arCode = "AR23" Or arCode = "AR22"
            If code = "1" Then
                ConvertValue = "Yes"
            Else
                ConvertValue = "No"
            End If
            
        ' Pattern 3: General Yes/No (any code with "Yes" in rule)
        Case InStr(conversionRule, "Yes") > 0
            If code = "1" Or code = "2" Or UCase(code) = "YES" Or UCase(code) = "Y" Then
                ConvertValue = "Yes"
            Else
                ConvertValue = "No"
            End If
            
        ' ================================================================
        ' NUMERIC CONVERSIONS - Handle percentages and decimals
        ' ================================================================
        Case InStr(conversionRule, "Percentage") > 0
            ' Convert to decimal if needed (e.g., 75% → 0.75 or keep as 75 depending on model)
            If IsNumeric(rawValue) Then
                ConvertValue = CDbl(rawValue)
            Else
                ConvertValue = ""
            End If
            
        ' ================================================================
        ' DATE CONVERSIONS - Ensure proper date formatting
        ' ================================================================
        Case InStr(conversionRule, "Date") > 0
            If IsDate(rawValue) Then
                ConvertValue = CDate(rawValue)
            Else
                ConvertValue = ""
            End If
            
        ' ================================================================
        ' DEFAULT - Pass through unconverted
        ' ================================================================
        Case Else
            ConvertValue = rawValue
    End Select
End Function

Function ColumnLetterToNumber(columnLetter As String) As Long
    ' Convert Excel column letter to number (e.g., "A" = 1, "AA" = 27, "AZ" = 52)
    ' Handles single and double letter columns
    
    On Error Resume Next
    ColumnLetterToNumber = Range(columnLetter & "1").Column
    If Err.Number <> 0 Then ColumnLetterToNumber = 0
    On Error GoTo 0
End Function

' ===================================================================
' UTILITY FUNCTIONS
' ===================================================================

Sub TestMapping()
    ' Quick test to verify mapping configuration
    ' Useful for troubleshooting without processing full loan tape
    
    Dim wsMapper As Worksheet
    Set wsMapper = ThisWorkbook.Sheets("BoE Auto-Mapper")
    
    Dim mappingCount As Long
    Dim criticalCount As Long
    Dim i As Long
    Dim arCode As String
    
    ' Count total mappings
    For i = 5 To wsMapper.Cells(wsMapper.Rows.Count, 1).End(xlUp).Row
        arCode = wsMapper.Cells(i, 1).Value
        If arCode <> "" And Left(arCode, 2) = "AR" Then
            mappingCount = mappingCount + 1
            
            ' Count critical fields
            If arCode = "AR3" Or arCode = "AR66" Or arCode = "AR67" Or _
               arCode = "AR131" Or arCode = "AR130" Or arCode = "AR141" Then
                criticalCount = criticalCount + 1
            End If
        End If
    Next i
    
    MsgBox "✓ Configuration Check Complete" & vbCrLf & vbCrLf & _
           "Total Mappings: " & mappingCount & vbCrLf & _
           "Critical Fields: " & criticalCount & " of 6" & vbCrLf & vbCrLf & _
           "Model Version: 5.2 - S&P Methodology Compliant" & vbCrLf & vbCrLf & _
           "Ready to process loan tapes with AR codes.", _
           vbInformation, "BoE Auto-Mapper Test"
End Sub

Sub ShowModelInfo()
    ' Display model version and key information
    
    MsgBox "S&P WAFF/WALS MODEL" & vbCrLf & vbCrLf & _
           "Version: 5.2" & vbCrLf & _
           "Status: S&P Methodology Compliant" & vbCrLf & _
           "Release: 18-Jan-2026" & vbCrLf & vbCrLf & _
           "KEY UPDATES IN v5.2:" & vbCrLf & _
           "• Help to Buy: 1.35x adjustment" & vbCrLf & _
           "• CCJ Tiers: 1.075x - 5.00x by age/value" & vbCrLf & _
           "• Arrears: 1.00x - 15.00x correct scale" & vbCrLf & _
           "• Geographic Concentration added" & vbCrLf & _
           "• Small Pool adjustment added" & vbCrLf & _
           "• Income Multiple joint thresholds" & vbCrLf & _
           "• Seasoning 9-tier scale" & vbCrLf & _
           "• Interest Only term-based tiers" & vbCrLf & vbCrLf & _
           "Based on S&P Global RMBS Methodology", _
           vbInformation, "Model Information"
End Sub

Sub ClearMappedData()
    ' Clear all data from Loan Tape (BoE) sheet
    ' Useful for starting fresh with new loan tape
    
    Dim response As VbMsgBoxResult
    response = MsgBox("This will clear ALL mapped data from 'Loan Tape (BoE)'." & vbCrLf & vbCrLf & _
                     "Are you sure you want to continue?", _
                     vbQuestion + vbYesNo, "Clear Mapped Data")
    
    If response = vbYes Then
        Dim wsBoE As Worksheet
        Set wsBoE = ThisWorkbook.Sheets("Loan Tape (BoE)")
        
        ' Clear from row 5 onwards (preserve headers)
        wsBoE.Range("A5:AZ" & wsBoE.Cells(wsBoE.Rows.Count, 1).End(xlUp).Row).ClearContents
        
        MsgBox "✓ Mapped data cleared successfully." & vbCrLf & vbCrLf & _
               "You can now map a new loan tape.", _
               vbInformation, "Data Cleared"
    End If
End Sub
