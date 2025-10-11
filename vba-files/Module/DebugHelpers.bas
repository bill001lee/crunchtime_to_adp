Attribute VB_Name = "DebugHelpers"
Option Explicit

' Global instance accessible from the Immediate window
Public p As PayrollProcessor

' Wire the processor to sheets in ThisWorkbook and register all outputs (incl. Rule 4)
Public Sub Debug_InitProcessor()
    Set p = New PayrollProcessor
    p.PublishAudit = True    ' Developer/audit mode
    ' Core inputs/lookups
    Set p.InputSheet = ThisWorkbook.Worksheets("DataIn")
    Set p.LookupSheet = ThisWorkbook.Worksheets("Lookup")
    Set p.ADPSheet = ThisWorkbook.Worksheets("ADP Pay Class")
    Set p.HolidaysSheet = ThisWorkbook.Worksheets("Holidays")
    Set p.ErrorSheet = ThisWorkbook.Worksheets("Errors")

    ' Outputs
    p.RegisterOutputSheet "Normal", ThisWorkbook.Worksheets("NormalTime")
    p.RegisterOutputSheet "OT_Rule1", ThisWorkbook.Worksheets("OTShiftHrs>5")
    p.RegisterOutputSheet "OT_Rule2", ThisWorkbook.Worksheets("OTDayHrs>11.5")
    p.RegisterOutputSheet "OT_Rule3", ThisWorkbook.Worksheets("OTWeekHrs>38")
    p.RegisterOutputSheet "OT_Rule4", ThisWorkbook.Worksheets("OTDays>5")
    p.RegisterOutputSheet "OT_Deduped", ThisWorkbook.Worksheets("OTDeduped")
    p.RegisterOutputSheet "Allowances", ThisWorkbook.Worksheets("AllowancesOut")

End Sub

' Inspect a single input row after extraction/lookup
Public Sub Debug_ProbeRow(Optional ByVal rowNum As Long = 2)
    Dim t As PayrollRowDataClass
    If p Is Nothing Then
        Debug.Print "Processor not initialized. Run Debug_InitProcessor first."
        Exit Sub
    End If

    Set t = p.Debug_GetRow(rowNum)
    If t Is Nothing Then
        Debug.Print "Row"; rowNum; ": extractor returned Nothing"
        Exit Sub
    End If

    Debug.Print "Row"; rowNum; _
        "; Emp=" & t.EmployeeCode & _
        "; Co=" & t.companyCode & _
        "; PayClass=" & t.PayClassCode & _
        "; PayCode=" & t.payrollCode & _
        "; Hours(units)=" & t.hours
End Sub

' Run the full pipeline and print counts per category
Public Sub Debug_RunAll()
    If p Is Nothing Then
        Debug.Print "Processor not initialized. Run Debug_InitProcessor first."
        Exit Sub
    End If

    p.ProcessPayroll

    Debug.Print "Normal rows:", ThisWorkbook.Worksheets("NormalTime").UsedRange.Rows.Count - 1
    Debug.Print "R1 rows:", ThisWorkbook.Worksheets("OTShiftHrs>5").UsedRange.Rows.Count - 1
    Debug.Print "R2 rows:", ThisWorkbook.Worksheets("OTDayHrs>11.5").UsedRange.Rows.Count - 1
    Debug.Print "Week>38 rows:", ThisWorkbook.Worksheets("OTWeekHrs>38").UsedRange.Rows.Count - 1
    Debug.Print "Days>5 rows:", ThisWorkbook.Worksheets("OTDays>5").UsedRange.Rows.Count - 1
    ' Note: Rule 4 is not produced by the full pipeline unless you add it there.
    If SheetExists("OTRule4") Then
        Debug.Print "R4 rows (full run):", ThisWorkbook.Worksheets("OTRule4").UsedRange.Rows.Count - 1
    End If
End Sub

' Utility
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function


