Attribute VB_Name = "BoE_Auto_Mapper"
Option Explicit

' ===================================================================
' BOE AUTO-MAPPER - AUTOMATIC FIELD MAPPING FOR S&P WAFF/WALS MODEL
' Version 4.0 - Updated for AR field codes
' ===================================================================

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
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.OnAction = "MapBoEFields"
    btn.Text = "Map BoE Fields"
    btn.Font.Bold = True
    btn.Font.Size = 11
    
    MsgBox "Button added! Click 'Map BoE Fields' after pasting loan tape.", vbInformation
End Sub

Sub MapBoEFields()
    Dim wsRaw As Worksheet, wsMapper As Worksheet, wsSP As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim boeCode As String, spColumn As String, conversionRule As String
    Dim colMapping As Object, headerRow As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Set wsRaw = ThisWorkbook.Sheets("Loan Tape (BoE Raw)")
    Set wsMapper = ThisWorkbook.Sheets("BoE Auto-Mapper")
    Set wsSP = ThisWorkbook.Sheets("Loan Tape (S&P)")
    
    ' Find header row (look for "AR" in first 20 rows)
    headerRow = 0
    For i = 1 To 20
        If InStr(1, wsRaw.Cells(i, 1).Value, "AR", vbTextCompare) > 0 Then
            headerRow = i
            Exit For
        End If
    Next i
    
    If headerRow = 0 Then
        MsgBox "Cannot find BoE field codes (AR1, AR2, etc.) in headers!" & vbCrLf & _
               "Please ensure your tape has AR codes.", vbCritical
        GoTo CleanUp
    End If
    
    ' Build mapping dictionary
    Set colMapping = CreateObject("Scripting.Dictionary")
    
    Dim mapperLastRow As Long
    mapperLastRow = wsMapper.Cells(wsMapper.Rows.Count, 1).End(xlUp).Row
    
    For i = 6 To mapperLastRow
        boeCode = wsMapper.Cells(i, 1).Value
        If boeCode <> "" Then
            spColumn = wsMapper.Cells(i, 3).Value
            conversionRule = wsMapper.Cells(i, 6).Value
            colMapping(boeCode) = Array(spColumn, conversionRule)
        End If
    Next i
    
    lastCol = wsRaw.Cells(headerRow, wsRaw.Columns.Count).End(xlToLeft).Column
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= headerRow Then
        MsgBox "No data found. Please paste your loan tape.", vbExclamation
        GoTo CleanUp
    End If
    
    wsSP.Range("A5:AZ" & wsSP.Cells(wsSP.Rows.Count, 1).End(xlUp).Row).ClearContents
    
    Dim dataCount As Long
    dataCount = 0
    
    For j = 1 To lastCol
        Dim headerText As String
        headerText = UCase(wsRaw.Cells(headerRow, j).Value)
        
        Dim extractedCode As String
        extractedCode = ExtractARCode(headerText)
        
        If extractedCode <> "" And colMapping.exists(extractedCode) Then
            spColumn = colMapping(extractedCode)(0)
            conversionRule = colMapping(extractedCode)(1)
            
            For i = headerRow + 1 To lastRow
                Dim rawValue As Variant, convertedValue As Variant
                rawValue = wsRaw.Cells(i, j).Value
                convertedValue = ConvertValue(rawValue, conversionRule)
                wsSP.Range(spColumn & (i - headerRow + 4)).Value = convertedValue
            Next i
            
            dataCount = dataCount + 1
        End If
    Next j
    
    MsgBox "Mapping complete!" & vbCrLf & vbCrLf & _
           "Fields mapped: " & dataCount & vbCrLf & _
           "Loans processed: " & (lastRow - headerRow), vbInformation
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub

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

Function ConvertValue(rawValue As Variant, conversionRule As String) As Variant
    If IsEmpty(rawValue) Or rawValue = "" Then
        ConvertValue = ""
        Exit Function
    End If
    
    If conversionRule = "Direct" Then
        ConvertValue = rawValue
        Exit Function
    End If
    
    Dim code As String
    code = CStr(rawValue)
    
    Select Case True
        Case InStr(conversionRule, "House") > 0
            Select Case code
                Case "1": ConvertValue = "House"
                Case "2": ConvertValue = "Flat"
                Case "3": ConvertValue = "Bungalow"
                Case "4": ConvertValue = "Maisonette"
                Case Else: ConvertValue = "Unknown"
            End Select
        Case InStr(conversionRule, "Fixed") > 0
            Select Case code
                Case "1": ConvertValue = "Fixed"
                Case "2": ConvertValue = "Variable"
                Case "3": ConvertValue = "Tracker"
                Case Else: ConvertValue = "Unknown"
            End Select
        Case InStr(conversionRule, "Owner Occupied") > 0
            Select Case code
                Case "1": ConvertValue = "Owner Occupied"
                Case "2": ConvertValue = "Buy to Let"
                Case "3": ConvertValue = "Second Home"
                Case Else: ConvertValue = "Unknown"
            End Select
        Case InStr(conversionRule, "Repayment") > 0
            Select Case code
                Case "1": ConvertValue = "Repayment"
                Case "2": ConvertValue = "Interest Only"
                Case "3": ConvertValue = "Part and Part"
                Case Else: ConvertValue = "Unknown"
            End Select
        Case InStr(conversionRule, "Yes") > 0
            If code = "1" Or code = "2" Then
                ConvertValue = "Yes"
            Else
                ConvertValue = "No"
            End If
        Case Else
            ConvertValue = rawValue
    End Select
End Function
