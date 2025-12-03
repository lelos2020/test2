Attribute VB_Name = "ValidationEngine"
Option Explicit

' ============================================================================
' Bank of England Loan Tape - Validation Engine
' ============================================================================
' This module contains the core validation logic for checking loan tape data
' against BoE requirements and cross-field business rules
' ============================================================================

Public Type ValidationError
    LoanRow As Long
    LoanID As String
    RuleID As String
    Category As String
    Severity As String
    Description As String
    FieldsInvolved As String
    CurrentValues As String
    RequiredAction As String
End Type

Public ErrorCollection() As ValidationError
Public ErrorCount As Long

' Main validation procedure - called from Dashboard button
Public Sub RunFullValidation()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startTime As Double
    Dim totalLoans As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    startTime = Timer
    
    ' Initialize
    ErrorCount = 0
    ReDim ErrorCollection(1 To 1000) ' Initial size, will resize if needed
    
    ' Clear previous error report
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    If ws.Cells(5, 1).Value <> "" Then
        ws.Range("A5:I" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).ClearContents
    End If
    
    ' Get data range
    Set ws = ThisWorkbook.Sheets("LoanData")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 5 Then
        MsgBox "No loan data found. Please enter loan data in the LoanData sheet.", vbInformation
        GoTo CleanUp
    End If
    
    totalLoans = lastRow - 4
    
    ' Update Dashboard status
    UpdateDashboardStatus "Running validation...", totalLoans
    
    ' Run validation categories
    ValidateMandatoryFields ws, lastRow
    ValidateRepaymentPaymentRules ws, lastRow
    ValidateBTLRules ws, lastRow
    ValidateInterestRateRules ws, lastRow
    ValidateBorrowerRules ws, lastRow
    ValidateArrearsRules ws, lastRow
    ValidateLTVRules ws, lastRow
    ValidateFurtherAdvances ws, lastRow
    ValidateCreditEnhancement ws, lastRow
    ValidateDateSequencing ws, lastRow
    
    ' Populate error report
    PopulateErrorReport
    
    ' Update dashboard with results
    UpdateDashboardResults totalLoans
    
    ' Show completion message
    Dim msg As String
    Dim elapsedTime As Double
    elapsedTime = Round(Timer - startTime, 2)
    
    If ErrorCount = 0 Then
        msg = "✓ VALIDATION COMPLETE" & vbCrLf & vbCrLf & _
              "All " & totalLoans & " loans passed validation!" & vbCrLf & _
              "Time: " & elapsedTime & " seconds"
        MsgBox msg, vbInformation, "Validation Successful"
    Else
        msg = "⚠ VALIDATION COMPLETE" & vbCrLf & vbCrLf & _
              "Found " & ErrorCount & " issues in " & totalLoans & " loans" & vbCrLf & _
              "Time: " & elapsedTime & " seconds" & vbCrLf & vbCrLf & _
              "Please review the ErrorReport sheet."
        MsgBox msg, vbExclamation, "Issues Found"
        
        ' Switch to error report
        ThisWorkbook.Sheets("ErrorReport").Activate
    End If
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' ============================================================================
' VALIDATION CATEGORY: Mandatory Fields
' ============================================================================
Private Sub ValidateMandatoryFields(ws As Worksheet, lastRow As Long)
    Dim i As Long, col As Long
    Dim fieldNum As String, loanID As String
    Dim mandatoryFields As Variant
    
    ' Define mandatory field columns (AR codes)
    mandatoryFields = Array("AR1", "AR2", "AR3", "AR5", "AR6", "AR7", "AR8", _
                           "AR19", "AR21", "AR26", "AR31", "AR43", "AR44", _
                           "AR55", "AR56", "AR59", "AR61", "AR66", "AR67", _
                           "AR69", "AR70", "AR71", "AR107", "AR109", "AR130", _
                           "AR131", "AR135", "AR136", "AR141", "AR143", "AR166")
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            fieldNum = ws.Cells(1, col).Value
            
            ' Check if this is a mandatory field
            If IsInArray(fieldNum, mandatoryFields) Then
                If IsEmpty(ws.Cells(i, col).Value) Or ws.Cells(i, col).Value = "" Then
                    LogError i, loanID, "MND", "Mandatory Fields", "CRITICAL", _
                             "Missing mandatory field: " & ws.Cells(2, col).Value, _
                             fieldNum, "Empty", _
                             "Mandatory field must be populated"
                End If
            End If
        Next col
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Repayment & Payment Rules
' ============================================================================
Private Sub ValidateRepaymentPaymentRules(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR69_col As Long, AR70_col As Long, AR71_col As Long
    Dim repaymentMethod As Variant, paymentFreq As Variant
    
    ' Find column positions
    AR69_col = FindColumnByAR(ws, "AR69") ' Repayment Method
    AR70_col = FindColumnByAR(ws, "AR70") ' Payment Frequency
    AR71_col = FindColumnByAR(ws, "AR71") ' Payment Due
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        repaymentMethod = ws.Cells(i, AR69_col).Value
        paymentFreq = ws.Cells(i, AR70_col).Value
        
        ' R1.1: Bullet repayment can only have Semi-Annual or Annual
        If repaymentMethod = 5 Then ' Bullet
            If paymentFreq <> 5 And paymentFreq <> 6 Then ' Not Semi-Annual or Annual
                LogError i, loanID, "R1.1", "Repayment & Payment", "CRITICAL", _
                         "Bullet repayment must have Semi-Annual or Annual payment frequency", _
                         "AR69, AR70", _
                         "AR69=" & repaymentMethod & ", AR70=" & paymentFreq, _
                         "Change AR70 to 5 (Semi-Annual) or 6 (Annual)"
            End If
        End If
        
        ' R1.2: Interest Only should have minimal payment
        If repaymentMethod = 2 Then ' Interest Only
            Dim paymentDue As Double
            If Not IsEmpty(ws.Cells(i, AR71_col).Value) Then
                paymentDue = ws.Cells(i, AR71_col).Value
                ' This is a warning, not critical
                If paymentDue > 0 Then
                    LogError i, loanID, "R1.2", "Repayment & Payment", "WARNING", _
                             "Interest Only loan should have payment approximately equal to interest", _
                             "AR69, AR71", _
                             "AR69=" & repaymentMethod & ", AR71=" & paymentDue, _
                             "Verify payment amount reflects interest only"
                End If
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Buy-to-Let Rules
' ============================================================================
Private Sub ValidateBTLRules(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR130_col As Long, AR154_col As Long, AR156_col As Long
    Dim occupancyType As String
    Dim rentalIncome As Variant, dscr As Variant
    
    ' Find columns
    AR130_col = FindColumnByAR(ws, "AR130") ' Occupancy Type
    AR154_col = FindColumnByAR(ws, "AR154") ' Estimated Rental Income
    AR156_col = FindColumnByAR(ws, "AR156") ' DSCR
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        occupancyType = Trim(UCase(ws.Cells(i, AR130_col).Value))
        
        If occupancyType = "BTL" Or InStr(occupancyType, "BUY-TO-LET") > 0 Then
            
            ' R2.1: BTL must have DSCR >= 100%
            If IsEmpty(ws.Cells(i, AR156_col).Value) Then
                LogError i, loanID, "R2.1", "BTL Validation", "CRITICAL", _
                         "BTL loan missing DSCR calculation", _
                         "AR130, AR156", _
                         "AR130=" & occupancyType & ", AR156=Empty", _
                         "Calculate and populate DSCR (must be >= 100%)"
            Else
                dscr = ws.Cells(i, AR156_col).Value
                If dscr < 100 Then
                    LogError i, loanID, "R2.1", "BTL Validation", "CRITICAL", _
                             "BTL DSCR below 100% indicates negative cash flow", _
                             "AR156", _
                             "DSCR=" & dscr & "%", _
                             "DSCR must be >= 100%. Review rental income and payment amounts."
                End If
            End If
            
            ' R2.2: BTL must have rental income
            If IsEmpty(ws.Cells(i, AR154_col).Value) Or ws.Cells(i, AR154_col).Value <= 0 Then
                LogError i, loanID, "R2.2", "BTL Validation", "CRITICAL", _
                         "BTL loan missing rental income", _
                         "AR130, AR154", _
                         "AR130=" & occupancyType & ", AR154=" & ws.Cells(i, AR154_col).Value, _
                         "BTL properties must have Estimated Rental Income > 0"
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Interest Rate Rules
' ============================================================================
Private Sub ValidateInterestRateRules(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR107_col As Long, AR108_col As Long, AR109_col As Long
    Dim AR110_col As Long, AR112_col As Long
    Dim rateType As Variant, indexRate As Variant, currentRate As Variant
    Dim margin As Variant, capRate As Variant
    
    ' Find columns
    AR107_col = FindColumnByAR(ws, "AR107") ' Interest Rate Type
    AR108_col = FindColumnByAR(ws, "AR108") ' Current Interest Rate Index
    AR109_col = FindColumnByAR(ws, "AR109") ' Current Interest Rate
    AR110_col = FindColumnByAR(ws, "AR110") ' Current Margin
    AR112_col = FindColumnByAR(ws, "AR112") ' Interest Cap Rate
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        If Not IsEmpty(ws.Cells(i, AR107_col).Value) Then
            rateType = ws.Cells(i, AR107_col).Value
            currentRate = ws.Cells(i, AR109_col).Value
            
            ' R3.1: Variable rate should equal Index + Margin
            If rateType = 1 Or rateType = 2 Then ' Floating/Variable
                If Not IsEmpty(ws.Cells(i, AR108_col).Value) And _
                   Not IsEmpty(ws.Cells(i, AR110_col).Value) And _
                   Not IsEmpty(currentRate) Then
                    
                    indexRate = ws.Cells(i, AR108_col).Value
                    margin = ws.Cells(i, AR110_col).Value
                    
                    If Abs(currentRate - (indexRate + margin)) > 0.25 Then
                        LogError i, loanID, "R3.1", "Interest Rate Rules", "WARNING", _
                                 "Current rate deviates from Index + Margin", _
                                 "AR107, AR108, AR109, AR110", _
                                 "Current=" & currentRate & "%, Index=" & indexRate & "%, Margin=" & margin & "%", _
                                 "Verify rate calculation. Tolerance: ±0.25%"
                    End If
                End If
            End If
            
            ' R3.2: Fixed rate should have zero margin
            If rateType = 3 Then ' Fixed
                If Not IsEmpty(ws.Cells(i, AR110_col).Value) Then
                    margin = ws.Cells(i, AR110_col).Value
                    If margin <> 0 Then
                        LogError i, loanID, "R3.2", "Interest Rate Rules", "WARNING", _
                                 "Fixed rate loan should have margin = 0", _
                                 "AR107, AR110", _
                                 "AR107=" & rateType & ", AR110=" & margin & "%", _
                                 "Set margin to 0 or blank for fixed rate loans"
                    End If
                End If
            End If
            
            ' R3.3: Capped rate must have cap rate populated and > current rate
            If rateType = 4 Then ' Capped
                If IsEmpty(ws.Cells(i, AR112_col).Value) Then
                    LogError i, loanID, "R3.3", "Interest Rate Rules", "CRITICAL", _
                             "Capped loan missing Interest Cap Rate", _
                             "AR107, AR112", _
                             "AR107=" & rateType & ", AR112=Empty", _
                             "Capped loans must have Interest Cap Rate populated"
                Else
                    capRate = ws.Cells(i, AR112_col).Value
                    If Not IsEmpty(currentRate) Then
                        If capRate <= currentRate Then
                            LogError i, loanID, "R3.3", "Interest Rate Rules", "CRITICAL", _
                                     "Interest Cap Rate must be greater than Current Rate", _
                                     "AR109, AR112", _
                                     "Current=" & currentRate & "%, Cap=" & capRate & "%", _
                                     "Cap rate must exceed current rate"
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Borrower Information Rules
' ============================================================================
Private Sub ValidateBorrowerRules(ws As Worksheet, lastRow As Long)
    Dim i As Long, col As Long
    Dim loanID As String
    Dim AR19_col As Long, AR26_col As Long, AR130_col As Long
    Dim numDebtors As Variant, primaryIncome As Variant
    Dim occupancyType As String
    
    ' Find columns
    AR19_col = FindColumnByAR(ws, "AR19") ' Number of Debtors
    AR26_col = FindColumnByAR(ws, "AR26") ' Primary Income
    AR130_col = FindColumnByAR(ws, "AR130") ' Occupancy Type
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        If Not IsEmpty(ws.Cells(i, AR19_col).Value) Then
            numDebtors = ws.Cells(i, AR19_col).Value
            
            ' R4.1: Single borrower - second borrower fields must be blank
            If numDebtors = 1 Then
                ' Check if any second borrower fields (AR189-AR210) are populated
                ' Simplified check - in production, check all second borrower fields
                For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                    Dim fieldNum As String
                    fieldNum = ws.Cells(1, col).Value
                    
                    If fieldNum Like "AR1[89]?" Or fieldNum Like "AR20?" Or fieldNum Like "AR21?" Then
                        If Not IsEmpty(ws.Cells(i, col).Value) And ws.Cells(i, col).Value <> "" Then
                            LogError i, loanID, "R4.1", "Borrower Information", "CRITICAL", _
                                     "Single borrower loan has second borrower data populated", _
                                     "AR19, " & fieldNum, _
                                     "AR19=" & numDebtors & ", " & fieldNum & "=" & ws.Cells(i, col).Value, _
                                     "Clear all second borrower fields when Number of Debtors = 1"
                            Exit For ' Only log once per loan
                        End If
                    End If
                Next col
            End If
        End If
        
        ' R4.3: Owner Occupied requires primary income > 0
        If Not IsEmpty(ws.Cells(i, AR130_col).Value) Then
            occupancyType = Trim(UCase(ws.Cells(i, AR130_col).Value))
            
            If InStr(occupancyType, "OWNER") > 0 Or InStr(occupancyType, "OCCUPIED") > 0 Then
                If IsEmpty(ws.Cells(i, AR26_col).Value) Or ws.Cells(i, AR26_col).Value <= 0 Then
                    LogError i, loanID, "R4.3", "Borrower Information", "CRITICAL", _
                             "Owner Occupied property missing primary income", _
                             "AR26, AR130", _
                             "AR130=" & occupancyType & ", AR26=" & ws.Cells(i, AR26_col).Value, _
                             "Owner Occupied properties must have Primary Income > 0"
                End If
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Arrears & Default Rules
' ============================================================================
Private Sub ValidateArrearsRules(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR166_col As Long, AR169_col As Long, AR170_col As Long
    Dim AR177_col As Long, AR178_col As Long
    Dim accountStatus As String, arrearsAmt As Variant, monthsArrears As Variant
    Dim defaultFlag As String, defaultDate As Variant
    
    ' Find columns
    AR166_col = FindColumnByAR(ws, "AR166") ' Account Status
    AR169_col = FindColumnByAR(ws, "AR169") ' Arrears Amount
    AR170_col = FindColumnByAR(ws, "AR170") ' Months in Arrears
    AR177_col = FindColumnByAR(ws, "AR177") ' Default Flag
    AR178_col = FindColumnByAR(ws, "AR178") ' Date of Default
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        monthsArrears = ws.Cells(i, AR170_col).Value
        arrearsAmt = ws.Cells(i, AR169_col).Value
        
        ' R5.1: Zero months in arrears requires zero arrears amount
        If monthsArrears = 0 Or IsEmpty(monthsArrears) Then
            If Not IsEmpty(arrearsAmt) And arrearsAmt > 0 Then
                LogError i, loanID, "R5.1", "Arrears & Default", "CRITICAL", _
                         "Arrears amount exists but months in arrears = 0", _
                         "AR169, AR170", _
                         "AR169=" & arrearsAmt & ", AR170=" & monthsArrears, _
                         "If Months in Arrears = 0, then Arrears Amount must = 0"
            End If
        End If
        
        ' R5.2: Arrears amount > 0 requires months > 0
        If Not IsEmpty(arrearsAmt) And arrearsAmt > 0 Then
            If IsEmpty(monthsArrears) Or monthsArrears = 0 Then
                LogError i, loanID, "R5.2", "Arrears & Default", "CRITICAL", _
                         "Arrears amount > 0 but months in arrears = 0", _
                         "AR169, AR170", _
                         "AR169=" & arrearsAmt & ", AR170=" & monthsArrears, _
                         "If Arrears Amount > 0, then Months in Arrears must > 0"
            End If
        End If
        
        ' R5.3: Account status must reflect arrears
        If Not IsEmpty(monthsArrears) And monthsArrears > 0 Then
            accountStatus = Trim(UCase(ws.Cells(i, AR166_col).Value))
            If InStr(accountStatus, "CURRENT") > 0 Then
                LogError i, loanID, "R5.3", "Arrears & Default", "CRITICAL", _
                         "Loan in arrears cannot have Account Status = Current", _
                         "AR166, AR170", _
                         "AR166=" & accountStatus & ", AR170=" & monthsArrears & " months", _
                         "Change status to Arrears, Default, or worse"
            End If
        End If
        
        ' R5.4: Defaulted loans must have default date
        If Not IsEmpty(ws.Cells(i, AR177_col).Value) Then
            defaultFlag = Trim(UCase(ws.Cells(i, AR177_col).Value))
            If defaultFlag = "Y" Or defaultFlag = "YES" Then
                If IsEmpty(ws.Cells(i, AR178_col).Value) Then
                    LogError i, loanID, "R5.4", "Arrears & Default", "CRITICAL", _
                             "Defaulted loan missing Date of Default", _
                             "AR177, AR178", _
                             "AR177=" & defaultFlag & ", AR178=Empty", _
                             "If Default Flag = Y, Date of Default is mandatory"
                End If
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: LTV Calculation Rules
' ============================================================================
Private Sub ValidateLTVRules(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR66_col As Long, AR67_col As Long, AR135_col As Long
    Dim AR136_col As Long, AR141_col As Long, AR143_col As Long
    Dim origBalance As Variant, origLTV As Variant, origValuation As Variant
    Dim currBalance As Variant, currLTV As Variant, currValuation As Variant
    Dim calcLTV As Double
    
    ' Find columns
    AR66_col = FindColumnByAR(ws, "AR66") ' Original Balance
    AR67_col = FindColumnByAR(ws, "AR67") ' Current Balance
    AR135_col = FindColumnByAR(ws, "AR135") ' Original LTV
    AR136_col = FindColumnByAR(ws, "AR136") ' Original Valuation
    AR141_col = FindColumnByAR(ws, "AR141") ' Current LTV
    AR143_col = FindColumnByAR(ws, "AR143") ' Current Valuation
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        ' R6.1: Original LTV calculation check
        origBalance = ws.Cells(i, AR66_col).Value
        origLTV = ws.Cells(i, AR135_col).Value
        origValuation = ws.Cells(i, AR136_col).Value
        
        If Not IsEmpty(origBalance) And Not IsEmpty(origLTV) And _
           Not IsEmpty(origValuation) And origValuation > 0 Then
            
            calcLTV = (origBalance / origValuation) * 100
            
            If Abs(calcLTV - origLTV) > 0.5 Then
                LogError i, loanID, "R6.1", "LTV Calculations", "WARNING", _
                         "Original LTV does not match calculation", _
                         "AR66, AR135, AR136", _
                         "Stated LTV=" & Format(origLTV, "0.00") & "%, Calculated=" & Format(calcLTV, "0.00") & "%", _
                         "Verify LTV calculation. Tolerance: ±0.5%"
            End If
        End If
        
        ' R6.2: Current LTV calculation check
        currBalance = ws.Cells(i, AR67_col).Value
        currLTV = ws.Cells(i, AR141_col).Value
        currValuation = ws.Cells(i, AR143_col).Value
        
        If Not IsEmpty(currBalance) And Not IsEmpty(currLTV) And _
           Not IsEmpty(currValuation) And currValuation > 0 Then
            
            calcLTV = (currBalance / currValuation) * 100
            
            If Abs(calcLTV - currLTV) > 0.5 Then
                LogError i, loanID, "R6.2", "LTV Calculations", "WARNING", _
                         "Current LTV does not match calculation", _
                         "AR67, AR141, AR143", _
                         "Stated LTV=" & Format(currLTV, "0.00") & "%, Calculated=" & Format(calcLTV, "0.00") & "%", _
                         "Verify LTV calculation. Tolerance: ±0.5%"
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Further Advances
' ============================================================================
Private Sub ValidateFurtherAdvances(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR88_col As Long, AR89_col As Long
    Dim furtherAdvAmt As Variant, furtherAdvDate As Variant
    
    ' Find columns
    AR88_col = FindColumnByAR(ws, "AR88") ' Further Loan Advance
    AR89_col = FindColumnByAR(ws, "AR89") ' Further Advance Date
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        furtherAdvAmt = ws.Cells(i, AR88_col).Value
        furtherAdvDate = ws.Cells(i, AR89_col).Value
        
        ' R7.1: Further advance amount requires date
        If Not IsEmpty(furtherAdvAmt) And furtherAdvAmt > 0 Then
            If IsEmpty(furtherAdvDate) Then
                LogError i, loanID, "R7.1", "Further Advances", "CRITICAL", _
                         "Further advance amount exists but date is missing", _
                         "AR88, AR89", _
                         "AR88=" & furtherAdvAmt & ", AR89=Empty", _
                         "If Further Advance Amount > 0, date must be populated"
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Credit Enhancement
' ============================================================================
Private Sub ValidateCreditEnhancement(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR78_col As Long, AR79_col As Long
    Dim migProvider As Variant, migAttachment As Variant
    
    ' Find columns
    AR78_col = FindColumnByAR(ws, "AR78") ' MIG Provider
    AR79_col = FindColumnByAR(ws, "AR79") ' MIG Attachment Point
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        migProvider = ws.Cells(i, AR78_col).Value
        migAttachment = ws.Cells(i, AR79_col).Value
        
        ' R8.2: MIG provider requires attachment point
        If Not IsEmpty(migProvider) And migProvider <> "" And migProvider <> "None" Then
            If IsEmpty(migAttachment) Or migAttachment = 0 Then
                LogError i, loanID, "R8.2", "Credit Enhancement", "CRITICAL", _
                         "MIG Provider specified but Attachment Point missing", _
                         "AR78, AR79", _
                         "AR78=" & migProvider & ", AR79=" & migAttachment, _
                         "If MIG Provider specified, Attachment Point must be populated"
            End If
        End If
    Next i
End Sub

' ============================================================================
' VALIDATION CATEGORY: Date Sequencing
' ============================================================================
Private Sub ValidateDateSequencing(ws As Worksheet, lastRow As Long)
    Dim i As Long
    Dim loanID As String
    Dim AR1_col As Long, AR55_col As Long, AR56_col As Long
    Dim cutoffDate As Date, origDate As Date, maturityDate As Date
    
    ' Find columns
    AR1_col = FindColumnByAR(ws, "AR1") ' Pool Cut-off Date
    AR55_col = FindColumnByAR(ws, "AR55") ' Loan Origination Date
    AR56_col = FindColumnByAR(ws, "AR56") ' Maturity Date
    
    For i = 5 To lastRow
        loanID = GetLoanID(ws, i)
        
        ' R9.1: Origination <= Cutoff < Maturity
        If IsDate(ws.Cells(i, AR1_col).Value) And _
           IsDate(ws.Cells(i, AR55_col).Value) And _
           IsDate(ws.Cells(i, AR56_col).Value) Then
            
            cutoffDate = ws.Cells(i, AR1_col).Value
            origDate = ws.Cells(i, AR55_col).Value
            maturityDate = ws.Cells(i, AR56_col).Value
            
            If Not (origDate <= cutoffDate And cutoffDate < maturityDate) Then
                LogError i, loanID, "R9.1", "Date Sequencing", "CRITICAL", _
                         "Date sequence violation: Origination <= Cutoff < Maturity", _
                         "AR1, AR55, AR56", _
                         "Orig=" & Format(origDate, "DD-MM-YYYY") & _
                         ", Cutoff=" & Format(cutoffDate, "DD-MM-YYYY") & _
                         ", Mat=" & Format(maturityDate, "DD-MM-YYYY"), _
                         "Dates must follow chronological sequence"
            End If
        End If
    Next i
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

' Log an error to the collection
Private Sub LogError(loanRow As Long, loanID As String, ruleID As String, _
                    Category As String, Severity As String, Description As String, _
                    FieldsInvolved As String, CurrentValues As String, RequiredAction As String)
    
    ErrorCount = ErrorCount + 1
    
    ' Resize array if needed
    If ErrorCount > UBound(ErrorCollection) Then
        ReDim Preserve ErrorCollection(1 To UBound(ErrorCollection) + 500)
    End If
    
    ' Store error details
    With ErrorCollection(ErrorCount)
        .LoanRow = loanRow
        .LoanID = loanID
        .RuleID = ruleID
        .Category = Category
        .Severity = Severity
        .Description = Description
        .FieldsInvolved = FieldsInvolved
        .CurrentValues = CurrentValues
        .RequiredAction = RequiredAction
    End With
End Sub

' Find column number for a given AR code
Private Function FindColumnByAR(ws As Worksheet, arCode As String) As Long
    Dim col As Long
    
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If UCase(Trim(ws.Cells(1, col).Value)) = UCase(arCode) Then
            FindColumnByAR = col
            Exit Function
        End If
    Next col
    
    FindColumnByAR = 0 ' Not found
End Function

' Get Loan ID from row (AR3 column)
Private Function GetLoanID(ws As Worksheet, Row As Long) As String
    Dim ar3Col As Long
    ar3Col = FindColumnByAR(ws, "AR3")
    
    If ar3Col > 0 Then
        GetLoanID = ws.Cells(Row, ar3Col).Value
    Else
        GetLoanID = "Row_" & Row
    End If
End Function

' Check if value is in array
Private Function IsInArray(val As Variant, arr As Variant) As Boolean
    Dim element As Variant
    
    For Each element In arr
        If element = val Then
            IsInArray = True
            Exit Function
        End If
    Next element
    
    IsInArray = False
End Function

' Populate error report sheet
Private Sub PopulateErrorReport()
    Dim ws As Worksheet
    Dim i As Long
    Dim rowNum As Long
    
    Set ws = ThisWorkbook.Sheets("ErrorReport")
    
    If ErrorCount = 0 Then Exit Sub
    
    ' Populate each error
    rowNum = 5 ' Start after header
    For i = 1 To ErrorCount
        With ErrorCollection(i)
            ws.Cells(rowNum, 1).Value = .LoanRow
            ws.Cells(rowNum, 2).Value = .LoanID
            ws.Cells(rowNum, 3).Value = .RuleID
            ws.Cells(rowNum, 4).Value = .Category
            ws.Cells(rowNum, 5).Value = .Severity
            ws.Cells(rowNum, 6).Value = .Description
            ws.Cells(rowNum, 7).Value = .FieldsInvolved
            ws.Cells(rowNum, 8).Value = .CurrentValues
            ws.Cells(rowNum, 9).Value = .RequiredAction
            
            ' Color code by severity
            Select Case .Severity
                Case "CRITICAL"
                    ws.Cells(rowNum, 5).Interior.Color = RGB(255, 199, 206)
                    ws.Cells(rowNum, 5).Font.Color = RGB(156, 0, 6)
                    ws.Cells(rowNum, 5).Font.Bold = True
                Case "WARNING"
                    ws.Cells(rowNum, 5).Interior.Color = RGB(255, 235, 156)
                    ws.Cells(rowNum, 5).Font.Color = RGB(156, 101, 0)
                    ws.Cells(rowNum, 5).Font.Bold = True
                Case "INFO"
                    ws.Cells(rowNum, 5).Interior.Color = RGB(198, 239, 206)
                    ws.Cells(rowNum, 5).Font.Color = RGB(0, 97, 0)
            End Select
        End With
        rowNum = rowNum + 1
    Next i
    
    ' Auto-fit columns
    ws.Columns("A:I").AutoFit
End Sub

' Update dashboard status during validation
Private Sub UpdateDashboardStatus(status As String, totalLoans As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ws.Range("B9").Value = totalLoans
    ws.Range("F12").Value = status
    ws.Range("F13").Value = Now
    
    DoEvents ' Allow screen to update
End Sub

' Update dashboard with final results
Private Sub UpdateDashboardResults(totalLoans As Long)
    Dim ws As Worksheet
    Dim criticalCount As Long, warningCount As Long, infoCount As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Count by severity
    For i = 1 To ErrorCount
        Select Case ErrorCollection(i).Severity
            Case "CRITICAL": criticalCount = criticalCount + 1
            Case "WARNING": warningCount = warningCount + 1
            Case "INFO": infoCount = infoCount + 1
        End Select
    Next i
    
    ' Update dashboard
    ws.Range("B9").Value = totalLoans
    ws.Range("B10").Value = totalLoans ' Validated
    ws.Range("B11").Value = criticalCount
    ws.Range("B12").Value = warningCount
    ws.Range("B13").Value = infoCount
    ws.Range("B14").Value = ErrorCount
    
    ' Update status
    If ErrorCount = 0 Then
        ws.Range("F12").Value = "✓ PASSED"
        ws.Range("F12").Interior.Color = RGB(198, 239, 206)
        ws.Range("F12").Font.Color = RGB(0, 97, 0)
    Else
        ws.Range("F12").Value = "⚠ ISSUES FOUND"
        ws.Range("F12").Interior.Color = RGB(255, 235, 156)
        ws.Range("F12").Font.Color = RGB(156, 101, 0)
    End If
    ws.Range("F12").Font.Bold = True
    
    ws.Range("F13").Value = Format(Now, "DD-MM-YYYY HH:MM")
    
    ' Update category breakdown (simplified - would need more detailed tracking)
    ' This is a placeholder - in full implementation, track errors by category
End Sub
