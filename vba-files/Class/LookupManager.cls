VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookupManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

+Option Explicit

' Worksheet references for lookups
Public LookupSheet As Worksheet
Public ADPSheet As Worksheet
Public HolidaysSheet As Worksheet

' Get company code from ownership entity

Public Function GetCompanyCode(ownershipEntity As Variant) As String
    Dim key As String, result As Variant
    key = Trim$(CStr(ownershipEntity))
    result = Application.VLookup(key, LookupSheet.Range("CompanyCode"), 2, False)
    If isError(result) Then
        GetCompanyCode = ""
    Else
        GetCompanyCode = CStr(result)
    End If
End Function


' Get pay class code based on pay rate and day of week
Public Function GetPayClassCode(payRate As Variant, dayOfWeek As Integer) As String
    If IsEmpty(payRate) Or payRate = 0 Then
        GetPayClassCode = "Y99"
        Exit Function
    End If
    
    Dim lookupRange As Range
    Set lookupRange = GetADPLookupRange(dayOfWeek, "PayClass")
    
    Dim result As Variant
    result = Application.VLookup(CDbl(Format(payRate, "0.00")), lookupRange, _
                                GetColumnOffset(dayOfWeek, "PayClass"), False)
    
    If isError(result) Then
        GetPayClassCode = "ERR"
    Else
        GetPayClassCode = CStr(result)
    End If
End Function

' Get cost centre based on pay rate and payroll export code
Public Function GetCostCentre(payRate As Variant, dayOfWeek As Integer, _
                              payrollExportCode As String) As String
    Dim suffix As Variant, lookupRange As Range, key As Double
    Set lookupRange = GetADPLookupRange(dayOfWeek, "Suffix")
    key = CDbl(Format(payRate, "0.00"))
    suffix = Application.VLookup(key, lookupRange, _
                                 GetColumnOffset(dayOfWeek, "Suffix"), False)
    If isError(suffix) Then suffix = ""
    GetCostCentre = CStr(suffix) & CStr(payrollExportCode)
End Function

' Get payroll code based on multiple factors
Public Function GetPayrollCode(payRate As Variant, dayOfWeek As Integer, _
                               payrollExportCode As String, DateIn As Date) As String
    Dim isHoliday As Boolean, lookupRange As Range, columnOffset As Integer
    Dim result As Variant, key As Double

    isHoliday = CheckIfHoliday(payrollExportCode, DateIn)
    key = CDbl(Format(payRate, "0.00"))

    If isHoliday Then
        Set lookupRange = GetADPLookupRange(dayOfWeek, "Holiday")
        columnOffset = GetColumnOffset(dayOfWeek, "Holiday")
    Else
        Set lookupRange = GetADPLookupRange(dayOfWeek, "Regular")
        columnOffset = GetColumnOffset(dayOfWeek, "Regular")
    End If

    result = Application.VLookup(key, lookupRange, columnOffset, False)
    If isError(result) Then
        GetPayrollCode = ""
    Else
        GetPayrollCode = CStr(result)
    End If
End Function

' Check if a date is a holiday
Private Function CheckIfHoliday(payrollExportCode As String, dateToCheck As Date) As Boolean
    Dim lookupKey As String
    lookupKey = payrollExportCode & Format(dateToCheck, "YYMMDD")
    
    Dim result As Variant
    result = Application.VLookup(lookupKey, HolidaysSheet.Range("A:B"), 2, False)
    
    CheckIfHoliday = Not isError(result)
End Function

' Get appropriate lookup range based on day of week
Private Function GetADPLookupRange(dayOfWeek As Integer, lookupType As String) As Range
    Select Case lookupType
        Case "PayClass"
            Select Case dayOfWeek
                Case 1 To 5: Set GetADPLookupRange = ADPSheet.Range("A:D")
                Case 6: Set GetADPLookupRange = ADPSheet.Range("B:D")
                Case 7: Set GetADPLookupRange = ADPSheet.Range("C:D")
            End Select
        Case "Suffix"
            Select Case dayOfWeek
                Case 1 To 5: Set GetADPLookupRange = ADPSheet.Range("A:J")
                Case 6: Set GetADPLookupRange = ADPSheet.Range("B:J")
                Case 7: Set GetADPLookupRange = ADPSheet.Range("C:J")
            End Select
        Case "Regular"
            Select Case dayOfWeek
                Case 1 To 5: Set GetADPLookupRange = ADPSheet.Range("A:I")
                Case 6: Set GetADPLookupRange = ADPSheet.Range("B:G")
                Case 7: Set GetADPLookupRange = ADPSheet.Range("C:H")
            End Select
        Case "Holiday"
            Select Case dayOfWeek
                Case 1 To 5: Set GetADPLookupRange = ADPSheet.Range("A:I")
                Case 6: Set GetADPLookupRange = ADPSheet.Range("B:I")
                Case 7: Set GetADPLookupRange = ADPSheet.Range("C:I")
            End Select
    End Select
End Function

' Get column offset for VLOOKUP based on day and type
Private Function GetColumnOffset(dayOfWeek As Integer, lookupType As String) As Integer
    Select Case lookupType
        Case "PayClass"
            Select Case dayOfWeek
                Case 1 To 5: GetColumnOffset = 4
                Case 6: GetColumnOffset = 3
                Case 7: GetColumnOffset = 2
            End Select
        Case "Suffix"
            Select Case dayOfWeek
                Case 1 To 5: GetColumnOffset = 10
                Case 6: GetColumnOffset = 9
                Case 7: GetColumnOffset = 8
            End Select
        Case "Regular"
            Select Case dayOfWeek
                Case 1 To 5: GetColumnOffset = 6
                Case 6: GetColumnOffset = 6
                Case 7: GetColumnOffset = 6
            End Select
        Case "Holiday"
            Select Case dayOfWeek
                Case 1 To 5: GetColumnOffset = 9
                Case 6: GetColumnOffset = 8
                Case 7: GetColumnOffset = 7
            End Select
    End Select
End Function

