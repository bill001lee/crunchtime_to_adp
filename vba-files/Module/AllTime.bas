Attribute VB_Name = "AllTime"
Option Explicit

' ==========================
' Public entry points (buttons)
' ==========================
Public Sub TranslatePayroll()
    On Error GoTo EH
    Dim prevCalc As XlCalculation, prevScreen As Boolean, prevEvents As Boolean
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim p As PayrollProcessor: Set p = InitializeAllProcessor(wb)

    If p Is Nothing Or Not p.IsInitialized Then
        MsgBox "Initialization failed. Check required sheet names.", vbCritical
        GoTo CleanExit
    End If

    p.ProcessPayroll

    ' Guard VersionInfo call
    On Error Resume Next
    VersionInfo.LogVersionToErrors p
    On Error GoTo EH

    MsgBox "Process complete: R1/R2/R3/R4 + OTDeduped + Normal + Allowances.", vbInformation

CleanExit:
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.Calculation = prevCalc
    Exit Sub

EH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' ==========================
' Initialization helpers
' ==========================
Private Function InitializeAllProcessor(Optional ByVal wb As Workbook) As PayrollProcessor
    On Error GoTo EH
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim p As New PayrollProcessor
    With p
        Set .InputSheet = GetRequiredWorksheet("DataIn", wb)
        Set .LookupSheet = GetRequiredWorksheet("Lookup", wb)
        Set .ADPSheet = GetRequiredWorksheet("ADP Pay Class", wb)
        Set .HolidaysSheet = GetRequiredWorksheet("Holidays", wb)
        .RegisterOutputSheet "Normal", GetRequiredWorksheet("NormalTime", wb)
        .RegisterOutputSheet "OT_Rule1", GetRequiredWorksheet("OTShiftHrs>5", wb)
        .RegisterOutputSheet "OT_Rule2", GetRequiredWorksheet("OTDayHrs>11.5", wb)
        .RegisterOutputSheet "OT_Rule3", GetRequiredWorksheet("OTWeekHrs>38", wb)
        .RegisterOutputSheet "OT_Rule4", GetRequiredWorksheet("OTDays>5", wb)
        .RegisterOutputSheet "OT_Deduped", GetRequiredWorksheet("OTDeduped", wb)
        .RegisterOutputSheet "Allowances", GetRequiredWorksheet("AllowancesOut", wb)
        Set .ErrorSheet = GetRequiredWorksheet("Errors", wb)
    End With
    p.PublishAudit = False   ' Production default: audit tabs disabled/hidden
    Set InitializeAllProcessor = p
    Exit Function
EH:
    Debug.Print "Err=" & Err.Number & ": " & Err.Description
    MsgBox "Required sheet missing: " & Err.Description, vbCritical
    Set InitializeAllProcessor = Nothing
End Function

Private Function GetRequiredWorksheet(sheetName As String, Optional ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set GetRequiredWorksheet = wb.Worksheets(sheetName)
    If GetRequiredWorksheet Is Nothing Then
        Dim anySheet As Object: Set anySheet = wb.Sheets(sheetName)
        If Not anySheet Is Nothing Then If TypeName(anySheet) = "Worksheet" Then Set GetRequiredWorksheet = anySheet
    End If
    On Error GoTo 0
    If GetRequiredWorksheet Is Nothing Then Err.Raise vbObjectError + 513, "AllTime", _
        "Required sheet '" & sheetName & "' is missing in '" & wb.name & "'."
End Function
'' Updated 2025-10-02 02:35

