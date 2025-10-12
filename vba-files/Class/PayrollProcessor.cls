VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PayrollProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ======= Class Module: PayrollProcessor =======
Option Explicit

' =========[ Public worksheet references ]=========
Public InputSheet As Worksheet
Public LookupSheet As Worksheet
Public ADPSheet As Worksheet
Public HolidaysSheet As Worksheet
Public ErrorSheet As Worksheet

Private sp As ShiftPieceClass
Private errorShiftKeys As Object  ' Dict: key -> True (shifts that errored in StoreWithCode)
Private outputSheets As Object ' Dict: categoryKey -> Worksheet
Private mPublishAudit As Boolean  ' default=False (production)

Public Property Get PublishAudit() As Boolean
    PublishAudit = mPublishAudit
End Property

Public Property Let PublishAudit(ByVal v As Boolean)
    mPublishAudit = v
End Property


' =========[ Public init check ]=========
Public Property Get IsInitialized() As Boolean
    Dim ok As Boolean
    ok = Not (InputSheet Is Nothing Or LookupSheet Is Nothing Or _
              ADPSheet Is Nothing Or HolidaysSheet Is Nothing)
    If Not ok Then
        IsInitialized = False
        Exit Property
    End If
    If outputSheets Is Nothing Or outputSheets.Count = 0 Then
        IsInitialized = False
        Exit Property
    End If
    If ErrorSheet Is Nothing Then
        IsInitialized = False
        Exit Property
    End If
    IsInitialized = True
End Property

' =========[ Registration of output sheets ]=========
Public Sub RegisterOutputSheet(sheetKey As String, ws As Worksheet)
    If outputSheets Is Nothing Then Set outputSheets = CreateObject("Scripting.Dictionary")
    Set outputSheets(sheetKey) = ws
End Sub

Public Sub ProcessPayroll()
    On Error GoTo EH

    If Not IsInitialized Then
        MsgBox "Processor not properly initialized.", vbCritical
        Exit Sub
    End If

    ' Init (once per run)
    Set errorShiftKeys = CreateObject("Scripting.Dictionary")

    ' Prepare outputs
    PrepareAllOutputSheets

    ' 1) Gather daily buckets
    Dim dayMap As Object
    GatherDailyBuckets dayMap

    ' 2) Optional audit views
    If PublishAudit Then
        Dim dictR1 As Object, dictR2 As Object, dictW3 As Object, dictR4 As Object
        Call ProcessRule1_Independent(dayMap, dictR1)
        Call ProcessRule2_DailyIndependent(dayMap, dictR2)
        Call ProcessRule3_WeeklyIndependent(dayMap, dictW3)
        Call ProcessRule4_DaysGT5Independent(dayMap, dictR4)
        WriteDictionaryToSheet dictR1, outputSheets("OT_Rule1")
        WriteDictionaryToSheet dictR2, outputSheets("OT_Rule2")
        WriteDictionaryToSheet dictW3, outputSheets("OT_Rule3")
        WriteDictionaryToSheet dictR4, outputSheets("OT_Rule4")
    End If
    SetAuditSheetVisibility PublishAudit

    ' *** IMPORTANT: export (PH / R4 / R2 / R3 / Base) FIRST ***
    ' This is where StoreWithCode runs and logs errors (context = StoreWithCode),
    ' populating 'errorShiftKeys' before allowances are generated.
    ProcessExportDeduped dayMap

    ' 3) Allowances (A100/A101) with suppression for error shifts
    ProcessAllowances

    ' 4) Rule 1 as Allowance (A102) with suppression for error shifts
    ProcessRule1_AsAllowance dayMap

    ' 5) Summary
    If PublishAudit Then
        LogError "Write", "", 0, "ProcessPayroll(Audit+Prod)", "Rows written", _
            "Normal=" & outputSheets("Normal").UsedRange.Rows.Count - 1 & _
            ", R1=" & outputSheets("OT_Rule1").UsedRange.Rows.Count - 1 & _
            ", R2=" & outputSheets("OT_Rule2").UsedRange.Rows.Count - 1 & _
            ", Week38=" & outputSheets("OT_Rule3").UsedRange.Rows.Count - 1 & _
            ", R4=" & outputSheets("OT_Rule4").UsedRange.Rows.Count - 1 & _
            ", OTDeduped=" & outputSheets("OT_Deduped").UsedRange.Rows.Count - 1 & _
            ", Allowances=" & outputSheets("Allowances").UsedRange.Rows.Count - 1
    Else
        LogError "Write", "", 0, "ProcessPayroll(Production)", "Rows written", _
            "Normal=" & outputSheets("Normal").UsedRange.Rows.Count - 1 & _
            ", OTDeduped=" & outputSheets("OT_Deduped").UsedRange.Rows.Count - 1 & _
            ", Allowances=" & outputSheets("Allowances").UsedRange.Rows.Count - 1
    End If

    Exit Sub
EH:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub SetAuditSheetVisibility(ByVal show As Boolean)
    Dim nm As Variant, k As String
    For Each nm In Array("OT_Rule1", "OT_Rule2", "OT_Rule3", "OT_Rule4")
        k = CStr(nm)
        If Not outputSheets Is Nothing Then
            If outputSheets.Exists(k) Then
                outputSheets(k).Visible = IIf(show, xlSheetVisible, xlSheetHidden)
            End If
        End If
    Next nm
End Sub

' =========[ Sheet preparation + Errors ]=========
Private Sub PrepareSingleOutputSheet(ByVal ws As Worksheet)
    With ws
        .Cells.Clear
        .Columns("A:H").NumberFormat = "@"
        Dim headers As Variant
        headers = Array( _
            "Company Code", "Employee Code", "Record Type", "Entry Date", _
            "Payroll Code", "Number of Hours", "Pay Class Code", "Cost Centre", _
            "From Date", "To Date", "Text", "Week Sort Key", "Date Sort Key" _
        )
        Dim c As Long
        For c = 0 To UBound(headers)
            .Cells(1, c + 1).value = headers(c)
        Next
        .Columns(2).NumberFormat = "@"
        .Columns(4).NumberFormat = "@"
        .Columns(5).NumberFormat = "@"
        .Columns(6).NumberFormat = "@"
        .Columns(9).NumberFormat = "@"
        .Columns(10).NumberFormat = "@"
        .Columns(12).NumberFormat = "0"
        .Columns(13).NumberFormat = "0"
    End With
End Sub

' ===== PATCH: Replace existing PrepareErrorsSheet =====
Private Sub PrepareErrorsSheet()
    With ErrorSheet
        .Cells.Clear
        ' A:Timestamp, B:Phase, C:Employee, D:Work Date, E:Hours (NEW), F:Context, G:Message, H:Details
        .Range("A1:H1").value = Array("Timestamp", "Phase", "Employee", "Work Date", "Hours", "Context", "Message", "Details")
        .Rows(1).Font.Bold = True
        .Columns.AutoFit
    End With
End Sub

' ===== PATCH: Replace existing LogError (keeps same signature) =====
Private Sub LogError(phase As String, emp As String, workDate As Date, ctx As String, msg As String, Optional details As String = "")
    If ErrorSheet Is Nothing Then Exit Sub
    On Error Resume Next
    With ErrorSheet
        Dim r As Long: r = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        .Cells(r, 1).value = Now           ' Timestamp
        .Cells(r, 2).value = phase         ' Phase
        .Cells(r, 3).value = emp           ' Employee
        .Cells(r, 4).value = workDate      ' Work Date
        .Cells(r, 5).value = ""            ' Hours (left blank in legacy calls)
        .Cells(r, 6).value = ctx           ' Context
        .Cells(r, 7).value = msg           ' Message
        .Cells(r, 8).value = details       ' Details
        .Columns.AutoFit
    End With
    On Error GoTo 0
End Sub

' ===== NEW: Explicit logger that writes Hours column =====
Private Sub LogErrorWithHours(phase As String, emp As String, workDate As Date, hours As Double, ctx As String, msg As String, Optional details As String = "")
    If ErrorSheet Is Nothing Then Exit Sub
    On Error Resume Next
    With ErrorSheet
        Dim r As Long: r = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        .Cells(r, 1).value = Now           ' Timestamp
        .Cells(r, 2).value = phase         ' Phase
        .Cells(r, 3).value = emp           ' Employee
        .Cells(r, 4).value = workDate      ' Work Date
        .Cells(r, 5).value = CLng(hours * 10000) ' Hours scaled like other outputs
        .Cells(r, 6).value = ctx           ' Context
        .Cells(r, 7).value = msg           ' Message
        .Cells(r, 8).value = details       ' Details
        .Columns.AutoFit
    End With
    On Error GoTo 0
End Sub

Private Sub PrepareAllOutputSheets()
    Dim k As Variant, isAudit As Boolean
    For Each k In outputSheets.keys
        isAudit = (k = "OT_Rule1" Or k = "OT_Rule2" Or k = "OT_Rule3" Or k = "OT_Rule4")
        If isAudit And Not PublishAudit Then
            ' Skip preparing audit sheets in production
        Else
            PrepareSingleOutputSheet outputSheets(k)
            ' If it's an audit sheet, add a top-level warning comment
            If isAudit And PublishAudit Then
                On Error Resume Next
                With outputSheets(k).Cells(1, 1)
                    .ClearComments
                    .AddComment "AUDIT VIEW ONLY  Do NOT export or sum with other OT tabs. " & _
                                "Use OTDeduped + NormalTime + AllowancesOut for payroll."
                End With
                On Error GoTo 0
            End If
        End If
    Next k
    PrepareErrorsSheet
End Sub


' Parse date from YYMMDD format
Public Function ParseYYMMDDDate(dateValue As Variant) As Date
    Dim s As String: s = Trim$(CStr(dateValue))
    s = Right$("000000" & s, 6)
    If Len(s) <> 6 Or Not IsNumeric(s) Then
        ParseYYMMDDDate = 0
        Exit Function
    End If
    Dim yy As Integer, mm As Integer, dd As Integer
    yy = 2000 + CInt(Mid$(s, 1, 2))
    mm = CInt(Mid$(s, 3, 2))
    dd = CInt(Mid$(s, 5, 2))
    ParseYYMMDDDate = DateSerial(yy, mm, dd)
End Function


' Format date from YYMMDD to DDMMYY

Public Function FormatDateForOutput(dateValue As Variant) As String
    FormatDateForOutput = Format(ParseYYMMDDDate(dateValue), "DDMMYY")
End Function


' Safe converter for YYMMDD variant to "DDMMYY" (preserving leading zeros)
Private Function YYMMDD_to_DDMMYY(ByVal v As Variant) As String
    Dim s As String: s = Right$("000000" & Trim$(CStr(v)), 6)
    YYMMDD_to_DDMMYY = Right$(s, 2) & Mid$(s, 3, 2) & Left$(s, 2)
End Function

' Validate numeric value
Public Function ValidateNumeric(value As Variant, defaultValue As Double) As Double
    If IsNumeric(value) Then
        ValidateNumeric = CDbl(value)
    Else
        ValidateNumeric = defaultValue
    End If
End Function

' Safe string concatenation with delimiter
Public Function SafeJoin(ParamArray values() As Variant) As String
    Dim result As String
    Dim i As Integer
    For i = 0 To UBound(values)
        If i > 0 Then result = result & vbCrLf
        result = result & CStr(values(i))
    Next i
    SafeJoin = result
End Function

Private Function ExtractShiftRaw(ByVal rowNum As Long) As PayrollRowDataClass
    Dim d As PayrollRowDataClass
    Set d = New PayrollRowDataClass
    On Error GoTo CleanExit

    With d
        ' Parse dates
        .DateIn = ParseYYMMDDandYYYYMMDDtoDate(InputSheet.Cells(rowNum, 7).value)
        .DateOut = ParseYYMMDDandYYYYMMDDtoDate(InputSheet.Cells(rowNum, 8).value)

        ' Time In/Out
        Dim vIn As Variant, vOut As Variant
        vIn = InputSheet.Cells(rowNum, 9).value
        vOut = InputSheet.Cells(rowNum, 10).value
        If IsDate(vIn) Then
            .DateTimeIn = .DateIn + CDbl(vIn)
        ElseIf IsNumeric(vIn) And vIn >= 0 And vIn < 1 Then
            .DateTimeIn = .DateIn + CDbl(vIn)
        Else
            .DateTimeIn = .DateIn
        End If
        If IsDate(vOut) Then
            .DateTimeOut = .DateOut + CDbl(vOut)
        ElseIf IsNumeric(vOut) And vOut >= 0 And vOut < 1 Then
            .DateTimeOut = .DateOut + CDbl(vOut)
        Else
            .DateTimeOut = .DateOut
        End If

        ' Overnight normalization
        If .DateTimeOut < .DateTimeIn Then
            .DateTimeOut = .DateTimeOut + 1
        End If

        ' Core fields
        .EmployeeCode = CStr(InputSheet.Cells(rowNum, 4).value)
        .payRate = InputSheet.Cells(rowNum, 11).value
        .payrollExportCode = CStr(InputSheet.Cells(rowNum, 2).value)
        .dayOfWeek = Weekday(.DateIn, vbMonday)

        
        ' Week ending date (robust)
        Dim we As Date
        we = ParseYYMMDDandYYYYMMDDtoDate(InputSheet.Cells(rowNum, 3).value)
        If we <> 0 Then
            .weekEnding = Format$(we, "DDMMYY")
        Else
            ' Fallback to tolerant 6-digit text parser
            .weekEnding = ParseYYMMDDtoDDMMYY(InputSheet.Cells(rowNum, 3).value)
        End If


        ' Entry/from/to dates
        .entryDate = ParseYYMMDDtoDDMMYY(InputSheet.Cells(rowNum, 3).value)
        .fromDate = ParseYYMMDDtoDDMMYY(InputSheet.Cells(rowNum, 7).value)
        .toDate = ParseYYMMDDtoDDMMYY(InputSheet.Cells(rowNum, 8).value)

        ' Sort keys
        .weekSortKey = CLng(Format(.DateIn - Weekday(.DateIn, vbMonday) + 1, "yyyymmdd"))
        .dateSortKey = CLng(Format$(.DateIn, "yyyymmdd"))

        ' Units: hours * 10000
        .hours = Round((.DateTimeOut - .DateTimeIn) * 24 * 10000, 0)
    End With

    ' Lookups
    Dim lm As New LookupManager
    Set lm.LookupSheet = LookupSheet
    Set lm.ADPSheet = ADPSheet
    Set lm.HolidaysSheet = HolidaysSheet
    d.companyCode = lm.GetCompanyCode(InputSheet.Cells(rowNum, 1).value)
    d.PayClassCode = lm.GetPayClassCode(d.payRate, d.dayOfWeek)
    d.costCentre = lm.GetCostCentre(d.payRate, d.dayOfWeek, d.payrollExportCode)
    If Len(Trim$(d.payrollCode)) = 0 Then
        d.payrollCode = lm.GetPayrollCode(d.payRate, d.dayOfWeek, d.payrollExportCode, d.DateIn)
    End If

CleanExit:
    Set ExtractShiftRaw = d
End Function

Private Function BuildAggregationKey(rowData As PayrollRowDataClass) As String
    Dim effectiveEntryDate As String
    Dim effectiveWeekSortKey As Long

    ' Always use week-ending for Entry Date in output
    effectiveEntryDate = rowData.weekEnding
    effectiveWeekSortKey = rowData.weekSortKey

    BuildAggregationKey = Join(Array( _
        rowData.companyCode, rowData.EmployeeCode, "E", effectiveEntryDate, _
        rowData.payrollCode, rowData.PayClassCode, rowData.costCentre, _
        rowData.fromDate, rowData.toDate, "", effectiveWeekSortKey, rowData.dateSortKey _
    ), vbCrLf)
End Function


Private Sub CopyPayrollRowData(ByRef src As PayrollRowDataClass, ByRef dest As PayrollRowDataClass)
    With dest
        .companyCode = src.companyCode
        .EmployeeCode = src.EmployeeCode

        ' >>> Ensure week-ending flows into derived rows
        .weekEnding = src.weekEnding

        ' Keep your original entryDate logic (not used for OTDeduped anymore, but harmless)
        .entryDate = IIf(IsOvertimeCode(src.payrollCode), src.fromDate, src.entryDate)

        .weekSortKey = IIf(IsOvertimeCode(src.payrollCode), _
            CLng(Format(ParseYYMMDDDate(src.fromDate) - Weekday(ParseYYMMDDDate(src.fromDate), vbMonday) + 1, "yyyymmdd")), _
            src.weekSortKey)

        .payrollCode = src.payrollCode
        .PayClassCode = src.PayClassCode
        .costCentre = src.costCentre
        .fromDate = src.fromDate
        .toDate = src.toDate
        .dateSortKey = src.dateSortKey
        .payrollExportCode = src.payrollExportCode
        .payRate = src.payRate
        .dayOfWeek = src.dayOfWeek
        .DateIn = src.DateIn
        .DateOut = src.DateOut
        .DateTimeIn = src.DateTimeIn
        .DateTimeOut = src.DateTimeOut
        .hours = src.hours
    End With
End Sub


Private Sub WriteOutputRowTo(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal key As String, ByVal hours As Double)
    Dim parts() As String
    parts = Split(key, vbCrLf)
    If UBound(parts) < 11 Then
        LogError "Write", "", 0, "WriteOutputRowTo", "Invalid key format", key
        Exit Sub
    End If

    With ws
        .Cells(rowNum, 1).value = parts(0) ' Company Code
        .Cells(rowNum, 2).value = parts(1) ' Employee Code
        .Cells(rowNum, 3).value = parts(2) ' Record Type
        .Cells(rowNum, 4).value = parts(3) ' Entry Date
        .Cells(rowNum, 5).value = parts(4) ' Payroll Code
        .Cells(rowNum, 6).value = hours    ' Number of Hours (units, hours*10000)
        .Cells(rowNum, 7).value = parts(5) ' Pay Class Code
        .Cells(rowNum, 8).value = parts(6) ' Cost Centre
        .Cells(rowNum, 9).value = parts(7) ' From Date
        .Cells(rowNum, 10).value = parts(8) ' To Date
        .Cells(rowNum, 11).value = parts(9) ' Text
        .Cells(rowNum, 12).value = parts(10) ' Week Sort Key
        .Cells(rowNum, 13).value = parts(11) ' Date Sort Key
    End With
End Sub

Private Sub WriteDictionaryToSheet(ByVal dict As Object, ByVal ws As Worksheet)
    Dim r As Long: r = 2
    Dim k As Variant
    For Each k In dict.keys
        WriteOutputRowTo ws, r, CStr(k), dict(k)
        r = r + 1
    Next
    ws.Columns.AutoFit
End Sub

Private Function dayKey(emp As String, daySortKey As Long) As String
    dayKey = emp & vbLf & CStr(daySortKey)
End Function

Private Sub GatherDailyBuckets(ByRef dayMap As Object)
    Set dayMap = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long, i As Long
    lastRow = InputSheet.Cells(InputSheet.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        Dim base As PayrollRowDataClass
        Set base = ExtractShiftRaw(i)
        If base.hours <= 0 Then GoTo ContinueLoop

        Dim r1Min As Double
        r1Min = ApplyOvertimeRule1(base.DateTimeIn, base.DateTimeOut)

        Dim sp As ShiftPieceClass
        Set sp = New ShiftPieceClass
        Set sp.data = base

        ' Copy PayrollRowData fields
        sp.companyCode = base.companyCode
        sp.EmployeeCode = base.EmployeeCode
        sp.entryDate = base.entryDate
        sp.payrollCode = base.payrollCode
        sp.PayClassCode = base.PayClassCode
        sp.costCentre = base.costCentre
        sp.fromDate = base.fromDate
        sp.toDate = base.toDate
        sp.weekSortKey = base.weekSortKey
        sp.dateSortKey = base.dateSortKey
        sp.payrollExportCode = base.payrollExportCode
        sp.payRate = base.payRate
        sp.dayOfWeek = base.dayOfWeek
        sp.DateIn = base.DateIn
        sp.DateOut = base.DateOut
        sp.DateTimeIn = base.DateTimeIn
        sp.DateTimeOut = base.DateTimeOut
        sp.hours = base.hours

        ' ShiftPiece fields
        sp.rawUnits = base.hours
        sp.r1Units = Round(r1Min / 60 * 10000, 0)
        sp.normalAvail = sp.rawUnits

        Dim k As String: k = dayKey(base.EmployeeCode, base.dateSortKey)
        If Not dayMap.Exists(k) Then dayMap.Add k, New Collection
        dayMap(k).Add sp

ContinueLoop:
    Next i
End Sub

'' --- Public Holiday check: site-aware, matches LookupManager.CheckIfHoliday ---
'Private Function IsPublicHolidayShift(ByRef sp As ShiftPieceClass) As Boolean
'    ' Uses HolidaysSheet A:B keyed by payrollExportCode & YYMMDD (same as LookupManager)
'    Dim key As String, v As Variant
'    If HolidaysSheet Is Nothing Then Exit Function
'    key = CStr(sp.payrollExportCode) & Format(sp.DateIn, "YYMMDD")
'    On Error Resume Next
'    v = Application.VLookup(key, HolidaysSheet.Range("A:B"), 2, False)
'    On Error GoTo 0
'    IsPublicHolidayShift = Not IsError(v)
'End Function

' --- Store PH units with code 603 (separate phase name for clean logging) ---
Private Sub StorePHPiece(ByRef dict As Object, ByRef d As PayrollRowDataClass, ByVal units As Double)
    Dim part As PayrollRowDataClass, why As String
    Set part = New PayrollRowDataClass
    CopyPayrollRowData d, part
    part.hours = units
    part.payrollCode = "603" ' fixed for Public Holiday
    If Not ValidateForNormalOrR1(part, why) Then
        LogError "PublicHoliday", part.EmployeeCode, part.DateIn, "StorePHPiece", why, _
                "Emp=" & part.EmployeeCode & ", Co=" & part.companyCode & _
                ", PayClass=" & part.PayClassCode & ", PayCode=" & part.payrollCode
        Exit Sub
    End If
    StoreByKey dict, part
End Sub
Private Function ValidateForNormalOrR1(ByRef d As PayrollRowDataClass, ByRef why As String) As Boolean
    If Len(Trim$(d.companyCode)) = 0 Then why = "Missing CompanyCode": Exit Function
    If Len(Trim$(d.payrollCode)) = 0 Then why = "Missing PayrollCode": Exit Function
    If UCase$(Trim$(d.PayClassCode)) = "ERR" Or Len(Trim$(d.PayClassCode)) = 0 Then
        why = "Invalid PayClassCode"
        Exit Function
    End If
    ValidateForNormalOrR1 = True
End Function

' Rule 2 now requires valid PayClass (from lookup) and will override PayrollCode only
Private Function ValidateForRule2(ByRef d As PayrollRowDataClass, ByRef why As String) As Boolean
    If Len(Trim$(d.companyCode)) = 0 Then why = "Missing CompanyCode": Exit Function
    ' Will be overridden, but must exist (we set it below)
    If UCase$(Trim$(d.PayClassCode)) = "ERR" Or Len(Trim$(d.PayClassCode)) = 0 Then
        why = "Invalid PayClassCode (lookup)"
        Exit Function
    End If
    ValidateForRule2 = True
End Function

' Replace your existing StoreByKey with this Variant-based version
Private Sub StoreByKey(ByRef dict As Object, ByRef rowData As PayrollRowDataClass)
    Dim key As String: key = BuildAggregationKey(rowData)
    If dict.Exists(key) Then
        dict(key) = CDbl(dict(key)) + rowData.hours
    Else
        dict.Add key, rowData.hours
    End If
End Sub

' Stores Rule 2 units (day > 11.5h), with day-based payroll code override
Private Sub StoreR2Piece(ByRef dict As Object, ByRef d As PayrollRowDataClass, units As Double, payrollCode As String)
    Dim part As PayrollRowDataClass
    Set part = New PayrollRowDataClass
    CopyPayrollRowData d, part
    part.hours = units
    part.payrollCode = payrollCode ' override for Rule 2

    Dim why As String
    If Not ValidateForRule2(part, why) Then
        LogError "Rule2", part.EmployeeCode, part.DateIn, "StoreR2Piece", why, _
            "Emp=" & part.EmployeeCode & ", Co=" & part.companyCode & _
            ", PayClass=" & part.PayClassCode & ", PayCode=" & part.payrollCode
        Exit Sub
    End If
    StoreByKey dict, part
End Sub

' Stores Rule 1 units (shift > 5h)
Private Sub StoreR1Piece(ByRef dict As Object, ByRef d As PayrollRowDataClass, units As Double)
    Dim part As PayrollRowDataClass
    Set part = New PayrollRowDataClass
    CopyPayrollRowData d, part
    part.hours = units
    part.payrollCode = "615" ' fixed for Rule 1

    Dim why As String
    If Not ValidateForNormalOrR1(part, why) Then
        LogError "Rule1", part.EmployeeCode, part.DateIn, "StoreR1Piece", why, _
            "Emp=" & part.EmployeeCode & ", Co=" & part.companyCode & _
            ", PayClass=" & part.PayClassCode & ", PayCode=" & part.payrollCode
        Exit Sub
    End If
    StoreByKey dict, part
End Sub

' Stores remaining normal units
Private Sub StoreNormalPiece(ByRef dict As Object, ByRef d As PayrollRowDataClass, units As Double)
    Dim part As PayrollRowDataClass
    Set part = New PayrollRowDataClass
    CopyPayrollRowData d, part
    part.hours = units

    ' === Log if overtime leaks into NormalTime ===
    If Not IsOvertimeCode(part.payrollCode) And units / 10000 > 11.5 Then
        If Not ErrorSheet Is Nothing Then
            Dim r As Long
            r = ErrorSheet.Cells(ErrorSheet.Rows.Count, "A").End(xlUp).Row + 1
            With ErrorSheet
                .Cells(r, 1).value = Now
                .Cells(r, 2).value = "NormalLeak"
                .Cells(r, 3).value = part.EmployeeCode
                .Cells(r, 4).value = part.DateIn
                .Cells(r, 5).value = "NormalTime > 11.5h"
                .Cells(r, 6).value = "PayRate=" & part.payRate
                .Cells(r, 7).value = "Units=" & Format(units / 10000, "0.0000")
            End With
        End If
    End If

    ' === Store the row ===
    StoreByKey dict, part
End Sub

Private Sub AllocateDailyOT(ByRef dayMap As Object, _
                            ByRef outNormal As Object, _
                            ByRef outR1 As Object, _
                            ByRef outR2 As Object)
    Set outNormal = CreateObject("Scripting.Dictionary")
    Set outR1 = CreateObject("Scripting.Dictionary")
    Set outR2 = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In dayMap.keys
        Dim shifts As Collection: Set shifts = dayMap(k)

        ' Total for the day (in units)
        Dim dayUnits As Double, i As Long
        For i = 1 To shifts.Count: dayUnits = dayUnits + shifts(i).rawUnits: Next

        ' Rule 2 (> 11.5h)
        Dim r2Units As Double: r2Units = Application.WorksheetFunction.Max(0, dayUnits - 11.5 * 10000)
        Dim first2 As Double: first2 = Application.WorksheetFunction.Min(20000, r2Units)
        Dim after2 As Double: after2 = Application.WorksheetFunction.Max(0, r2Units - 20000)

        ' Get PayrollCode mapping for Rule 2
        Dim code1 As String, code2 As String, tmpA As Double, tmpB As Double
        Call ApplyOvertimeRule2( _
            shifts(1).data.DateTimeIn, shifts(1).data.DateTimeIn + (12# / 24#), _
            shifts(1).data.dayOfWeek, _
            code1, code2, tmpA, tmpB)

        ' Allocate R2 from end of day backwards: AFTER 2h first (code2), then FIRST 2h (code1)
        Dim need As Double, take As Double

        ' AFTER 2 hours (code2)
        need = after2
        For i = shifts.Count To 1 Step -1
            take = Application.WorksheetFunction.Min(need, shifts(i).normalAvail)
            If take > 0 Then
                StoreR2Piece outR2, shifts(i).data, take, code2
                shifts(i).normalAvail = shifts(i).normalAvail - take
                need = need - take
                If need <= 0 Then Exit For
            End If
        Next i

        ' FIRST 2 hours (code1)
        need = first2
        For i = shifts.Count To 1 Step -1
            take = Application.WorksheetFunction.Min(need, shifts(i).normalAvail)
            If take > 0 Then
                StoreR2Piece outR2, shifts(i).data, take, code1
                shifts(i).normalAvail = shifts(i).normalAvail - take
                need = need - take
                If need <= 0 Then Exit For
            End If
        Next i

        ' Now Rule 1 on remaining portions (no double counting with R2)
        For i = 1 To shifts.Count
            Dim r1Take As Double
            r1Take = Application.WorksheetFunction.Min(shifts(i).r1Units, shifts(i).normalAvail)
            If r1Take > 0 Then
                StoreR1Piece outR1, shifts(i).data, r1Take
                shifts(i).normalAvail = shifts(i).normalAvail - r1Take
            End If
            ' Remaining is Normal
            If shifts(i).normalAvail > 0 Then
                StoreNormalPiece outNormal, shifts(i).data, shifts(i).normalAvail
            End If
        Next i
    Next k
End Sub

' =========[ Weekly rules ]=========
Private Function weekStart(d As Date) As Date
    weekStart = d - (Weekday(d, vbMonday) - 1)
End Function

Private Function weekKey(emp As String, d As Date) As String
    weekKey = emp & vbLf & Format(weekStart(d), "yyyymmdd")
End Function

Private Function weekKey2(emp As String, w As Long) As String
    weekKey2 = emp & "|" & CStr(w)
End Function

' Convert dictionary entries to a list for weekly processing
Private Sub ExpandDictToList(ByVal dict As Object, ByRef listOut As Collection)
    Set listOut = New Collection
    Dim k As Variant
    For Each k In dict.keys
        Dim parts() As String: parts = Split(CStr(k), vbCrLf)
        Dim item As PayrollRowDataClass
        Set item = New PayrollRowDataClass
        item.companyCode = parts(0)
        item.EmployeeCode = parts(1)
        item.entryDate = parts(3)
        item.payrollCode = parts(4)
        item.PayClassCode = parts(5)
        item.costCentre = parts(6)
        item.fromDate = parts(7)
        item.toDate = parts(8)
        item.weekSortKey = CLng(parts(10))
        item.dateSortKey = CLng(parts(11))
        ' Rebuild DateIn from dateSortKey
        Dim ds As String: ds = Right$("00000000" & CStr(item.dateSortKey), 8)
        item.DateIn = DateSerial(CInt(Left$(ds, 4)), CInt(Mid$(ds, 5, 2)), CInt(Right$(ds, 2)))
        item.dayOfWeek = Weekday(item.DateIn, vbMonday)
        item.hours = dict(k)
        listOut.Add item
    Next k
End Sub

' Weekly rule: Days > 5 (convert entire 6th/7th worked days)
Private Sub ApplyWeeklyDaysGT5(ByRef normals As Collection, ByRef outDaysGT5 As Object)
    Set outDaysGT5 = CreateObject("Scripting.Dictionary")

    ' Build (emp, weekSortKey) -> dayKey(yyyymmdd) -> list of indexes
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To normals.Count
        Dim it As PayrollRowDataClass
        Set it = normals(i)
        If it.hours <= 0 Then GoTo NextI

        Dim wkKey As String: wkKey = weekKey2(it.EmployeeCode, CLng(it.weekSortKey))  ' <- standardized
        If Not map.Exists(wkKey) Then map(wkKey) = CreateObject("Scripting.Dictionary")

        Dim dayKeyStr As String: dayKeyStr = Format(it.DateIn, "yyyymmdd")            ' use actual calendar day
        If Not map(wkKey).Exists(dayKeyStr) Then map(wkKey)(dayKeyStr) = New Collection
        map(wkKey)(dayKeyStr).Add i
NextI:
    Next i

    ' For each (emp, weekSortKey), order worked days chronologically and convert day #6 and #7
    Dim wk As Variant
    For Each wk In map.keys
        ' Gather the day keys and sort ascending numerically
        Dim daysDict As Object: Set daysDict = map(wk)
        Dim arr() As String, c As Long, ds As Variant
        ReDim arr(1 To daysDict.Count)
        c = 0
        For Each ds In daysDict.keys
            c = c + 1
            arr(c) = CStr(ds)
        Next ds

        Dim a As Long, b As Long, tmp As String
        For a = 1 To UBound(arr) - 1
            For b = a + 1 To UBound(arr)
                If CLng(arr(a)) > CLng(arr(b)) Then tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
            Next b
        Next a

        ' Walk ordered days and convert 6th/7th days
        Dim workedCount As Long: workedCount = 0
        For a = 1 To UBound(arr)
            workedCount = workedCount + 1
            If workedCount >= 6 Then
                Dim col As Collection: Set col = daysDict(arr(a))
                Dim j As Long
                For j = 1 To col.Count
                    Dim idx As Long: idx = CLng(col(j))
                    Dim seg As PayrollRowDataClass
                    Set seg = normals(idx)
                    If seg.hours > 0 Then
                        Dim ov As String: ov = GetPayrollCode_DaysGT5(seg.dayOfWeek) ' "" keeps lookup-derived code
                        If Len(ov) > 0 Then seg.payrollCode = ov
                        StoreByKey outDaysGT5, seg
                        seg.hours = 0
                    End If
                Next j
            End If
        Next a
    Next wk

End Sub

Private Sub ApplyWeeklyHoursGT38_FromRaw(ByRef dayMap As Object, ByRef outWeekGT38 As Object)
    Set outWeekGT38 = CreateObject("Scripting.Dictionary")
    Dim workDate As Date

    ' Group shifts by (employee, weekSortKey)
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim k As Variant, shifts As Collection, i As Long
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        For i = 1 To shifts.Count
            Dim wkKey As String: wkKey = weekKey2(shifts(i).EmployeeCode, shifts(i).weekSortKey)
            workDate = IIf(IsDate(shifts(i).DateIn), shifts(i).DateIn, Date)
            If Not map.Exists(wkKey) Then map.Add wkKey, New Collection
            map(wkKey).Add shifts(i)
        Next i
    Next k

    ' Process each week
    Dim wk As Variant
    For Each wk In map.keys
        Dim totalUnits As Double: totalUnits = 0
        Dim shiftList As Collection: Set shiftList = map(wk)
        For i = 1 To shiftList.Count
            totalUnits = totalUnits + shiftList(i).rawUnits
        Next i

        Dim excessUnits As Double
        Dim need As Double
        excessUnits = totalUnits - (38# * 10000)
        need = Application.WorksheetFunction.Max(0, excessUnits)

        Dim empCode As String: empCode = Split(wk, "|")(0)
        Dim entryDate As String: entryDate = Split(wk, "|")(1)
        workDate = IIf(IsDate(shiftList(1).DateIn), shiftList(1).DateIn, Date)

        LogError "Rule3", empCode, workDate, "AllocationStart", "Checking weekly total", _
            "entryDate=" & entryDate & ", TotalUnits=" & totalUnits & ", Need=" & need

        If need <= 0 Then GoTo NextWeek3

        ' Sort shifts by date descending
        Dim sorted() As Variant: ReDim sorted(1 To shiftList.Count)
        For i = 1 To shiftList.Count: Set sorted(i) = shiftList(i): Next i
        Dim a As Long, b As Long
        For a = 1 To UBound(sorted) - 1
            For b = a + 1 To UBound(sorted)
                If sorted(a).dateSortKey < sorted(b).dateSortKey Then
                    Dim tmp As Variant: Set tmp = sorted(a): Set sorted(a) = sorted(b): Set sorted(b) = tmp
                End If
            Next b
        Next a

        ' Allocate overtime from latest day backward
        For i = 1 To UBound(sorted)
            Dim take As Double: take = Application.WorksheetFunction.Min(need, sorted(i).rawUnits)
            If take > 0 Then
                workDate = IIf(IsDate(sorted(i).DateIn), sorted(i).DateIn, Date)
                LogError "Rule3", empCode, workDate, "Allocating", "Taking overtime", _
                    "Take=" & take & ", rawUnits=" & sorted(i).rawUnits

                Dim part As PayrollRowDataClass: Set part = New PayrollRowDataClass
                CopyPayrollRowData sorted(i).data, part
                part.hours = take
                Dim ov As String: ov = GetPayrollCode_WeekGT38(part.dayOfWeek)
                If Len(ov) > 0 Then part.payrollCode = ov
                StoreByKey outWeekGT38, part
                need = need - take

                LogError "Rule3", empCode, workDate, "PostAlloc", "Remaining overtime need", _
                    "RemainingNeed=" & need

                If need <= 0 Then Exit For
            End If
        Next i

NextWeek3:
    Next wk
End Sub

' Weekly rule: Week > 38 (allocate from end of week backwards)  standardized by weekSortKey
Private Sub ApplyWeeklyHoursGT38(ByRef normals As Collection, ByRef outWeekGT38 As Object)
    Set outWeekGT38 = CreateObject("Scripting.Dictionary")

    ' Group indexes per (emp, weekSortKey)
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To normals.Count
        Dim it As PayrollRowDataClass
        Set it = normals(i)
        If it.hours <= 0 Then GoTo NextI

        Dim wkKey As String: wkKey = weekKey2(it.EmployeeCode, CLng(it.weekSortKey))  ' <- standardized
        If Not map.Exists(wkKey) Then map(wkKey) = New Collection
        map(wkKey).Add i
NextI:
    Next i

    Dim wk As Variant
    For Each wk In map.keys
        ' Compute total remaining normal units in the week
        Dim total As Double, j As Long
        For j = 1 To map(wk).Count
            Dim idxTotal As Long: idxTotal = map(wk)(j)
            total = total + normals(idxTotal).hours
        Next j

        Dim need As Double: need = Application.WorksheetFunction.Max(0, total - 38 * 10000)
        If need <= 0 Then GoTo NextWeek

        ' Sort weekly indexes by date ascending to allocate from end
        Dim arr() As Long: ReDim arr(1 To map(wk).Count)
        For j = 1 To UBound(arr): arr(j) = map(wk)(j): Next j

        Dim a As Long, b As Long, tmp As Long
        For a = 1 To UBound(arr) - 1
            For b = a + 1 To UBound(arr)
                If normals(arr(a)).dateSortKey > normals(arr(b)).dateSortKey Then
                    tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
                End If
            Next b
        Next a

        ' Allocate from the latest day backwards
        For j = UBound(arr) To 1 Step -1
            Dim idx As Long: idx = arr(j)
            Dim seg As PayrollRowDataClass
            Set seg = normals(idx)
            If seg.hours <= 0 Then GoTo NextJ

            Dim take As Double: take = Application.WorksheetFunction.Min(need, seg.hours)
            If take > 0 Then
                Dim part As PayrollRowDataClass
                Set part = New PayrollRowDataClass
                CopyPayrollRowData seg, part
                part.hours = take

                Dim ov As String: ov = GetPayrollCode_WeekGT38(part.dayOfWeek) ' "" keeps lookup-derived code
                If Len(ov) > 0 Then part.payrollCode = ov

                StoreByKey outWeekGT38, part
                seg.hours = seg.hours - take
                need = need - take
                If need <= 0 Then Exit For
            End If
NextJ:
        Next j
NextWeek:
    Next wk
End Sub

' =========[ Weekly override hooks (return "" to keep lookup) ]=========
Private Function GetPayrollCode_WeekGT38(dayOfWeek As Integer) As String
    GetPayrollCode_WeekGT38 = "" ' set later if you decide to override
End Function

' =========[ Weekly override hooks (Days>5 uses code 610) ]=========
Private Function GetPayrollCode_DaysGT5(dayOfWeek As Integer) As String
    ' Force payroll code 610 for the 6th and 7th worked days in a week
    GetPayrollCode_DaysGT5 = "610"
End Function
' =========[ Existing OT rule helpers (kept for reuse/compatibility) ]=========
' Rule 1: minutes beyond 5h in a single shift; also sets payrollCode when ByRef provided
Private Function ApplyOvertimeRule1( _
    DateTimeIn As Date, _
    DateTimeOut As Date, _
    Optional ByRef payrollCode As String = "" _
) As Double
    payrollCode = "615"
    Dim totalMinutes As Double
    totalMinutes = (DateTimeOut - DateTimeIn) * 24 * 60
    If totalMinutes > 300 Then
        ApplyOvertimeRule1 = totalMinutes - 300
    Else
        ApplyOvertimeRule1 = 0
    End If
End Function

Private Function ApplyOvertimeRule2( _
    DateTimeIn As Date, _
    DateTimeOut As Date, _
    dayOfWeek As Integer, _
    ByRef code1 As String, _
    ByRef code2 As String, _
    ByRef firstTwoHours As Double, _
    ByRef afterTwoHours As Double _
) As Boolean
    code1 = "": code2 = ""
    firstTwoHours = 0: afterTwoHours = 0
    Select Case dayOfWeek ' Weekday(..., vbMonday): Mon=1,...,Sun=7
        Case 1 To 5: code1 = "605": code2 = "606"
        Case 6:      code1 = "607": code2 = "608"  ' <-- FIXED
        Case 7: code1 = "602": code2 = "602"
    End Select
    ' ... (rest unchanged) ...
End Function
' Fail-fast getter inside the class (so we can register outputs from legacy path)
Private Function GetRequiredWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetRequiredWorksheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If GetRequiredWorksheet Is Nothing Then
        Err.Raise vbObjectError + 513, "PayrollProcessor", _
            "Required sheet '" & sheetName & "' is missing."
    End If
End Function

Private Function ParseYYMMDDandYYYYMMDDtoDate(ByVal v As Variant) As Date
    ' Accepts:
    '   - Excel Date/serial (IsDate = True)
    '   - "yyyymmdd" (8 digits)
    '   - "yymmdd"   (6 digits, year = 2000 + yy)
    '   - "yyyy-mm-dd" / "yyyy/mm/dd" (and similar with separators)
    ' Returns 0 if it cannot parse.

    Dim d As Date
    Dim s As String

    ' 1) If Excel already sees a valid Date/serial, take it
    If IsDate(v) Then
        ParseYYMMDDandYYYYMMDDtoDate = CDate(v)
        Exit Function
    End If

    ' 2) Normalize to digits-only string when it looks like a date text
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    ' remove common separators
    s = Replace(Replace(s, "-", ""), "/", "")
    s = Replace(Replace(s, " ", ""), ".", "")

    ' 3) 8-digit: yyyymmdd
    If Len(s) = 8 And IsNumeric(s) Then
        ParseYYMMDDandYYYYMMDDtoDate = DateSerial( _
            CInt(Left$(s, 4)), CInt(Mid$(s, 5, 2)), CInt(Right$(s, 2)) _
        )
        Exit Function
    End If

    ' 4) 6-digit: yymmdd  -> 2000 + yy
    If Len(s) = 6 And IsNumeric(s) Then
        ParseYYMMDDandYYYYMMDDtoDate = DateSerial( _
            2000 + CInt(Left$(s, 2)), CInt(Mid$(s, 3, 2)), CInt(Right$(s, 2)) _
        )
        Exit Function
    End If

    ' 5) As a last resort, if it's numeric (likely a serial), try CDate
    If IsNumeric(v) Then
        On Error Resume Next
        d = CDate(v)
        On Error GoTo 0
        If d <> 0 Then
            ParseYYMMDDandYYYYMMDDtoDate = d
            Exit Function
        End If
    End If

    ' 6) Could not parse; return 0 (midnight) and your caller will log later
    ParseYYMMDDandYYYYMMDDtoDate = 0
End Function


Private Function ParseYYMMDDtoDDMMYY(ByVal v As Variant) As String
    ' Accepts 6-digit YYMMDD (text or number). Returns DDMMYY.
    ' Tolerates separators and pads leading zeros.
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function

    ' Strip common separators/spaces
    s = Replace(s, "-", "")
    s = Replace(s, "/", "")
    s = Replace(s, " ", "")

    ' Coerce to 6 digits with leading zeros (handles numbers like 60125 -> 060125)
    If Not IsNumeric(s) Then Exit Function
    s = Right$("000000" & s, 6)

    ParseYYMMDDtoDDMMYY = Right$(s, 2) & Mid$(s, 3, 2) & Left$(s, 2)
End Function

' Debug helper: allows Immediate Window to call the private extractor
Public Function Debug_GetRow(ByVal rowNum As Long) As PayrollRowDataClass
    Set Debug_GetRow = ExtractShiftRaw(rowNum)
End Function

'Public Sub ProcessRule2Only()
'    If Not IsInitialized Then
'        MsgBox "Processor not properly initialized.", vbCritical
'        Exit Sub
'    End If
'
'    'PrepareAllOutputSheets
'
'    Dim dayMap As Object
'    GatherDailyBuckets dayMap
'
'    Dim dictNormal As Object: Set dictNormal = CreateObject("Scripting.Dictionary")
'    Dim dictR2 As Object: Set dictR2 = CreateObject("Scripting.Dictionary")
'
'    Dim k As Variant
'    For Each k In dayMap.Keys
'        Dim shifts As Collection: Set shifts = dayMap(k)
'
'        ' Reset dayUnits for each new day
'        Dim dayUnits As Double: dayUnits = 0
'
'        ' Accumulate total units for the current day
'        Dim i As Long
'        For i = 1 To shifts.Count
'            dayUnits = dayUnits + shifts(i).rawUnits
'        Next
'
'        If dayUnits > 11.5 * 10000 Then
'            Dim r2Units As Double: r2Units = dayUnits - 11.5 * 10000
'            Dim first2 As Double: first2 = Application.WorksheetFunction.Min(20000, r2Units)
'            Dim after2 As Double: after2 = Application.WorksheetFunction.Max(0, r2Units - 20000)
'            Dim code1 As String, code2 As String, tmpA As Double, tmpB As Double
'            Call ApplyOvertimeRule2(shifts(1).data.DateTimeIn, shifts(1).data.DateTimeIn + (12# / 24#), shifts(1).data.dayOfWeek, code1, code2, tmpA, tmpB)
'
'            Dim need As Double, take As Double
'            need = after2
'            For i = shifts.Count To 1 Step -1
'                take = Application.WorksheetFunction.Min(need, shifts(i).normalAvail)
'                If take > 0 Then
'                    StoreR2Piece dictR2, shifts(i).data, take, code2
'                    shifts(i).normalAvail = shifts(i).normalAvail - take
'                    need = need - take
'                    If need <= 0 Then Exit For
'                End If
'            Next i
'
'            need = first2
'            For i = shifts.Count To 1 Step -1
'                take = Application.WorksheetFunction.Min(need, shifts(i).normalAvail)
'                If take > 0 Then
'                    StoreR2Piece dictR2, shifts(i).data, take, code1
'                    shifts(i).normalAvail = shifts(i).normalAvail - take
'                    need = need - take
'                    If need <= 0 Then Exit For
'                End If
'            Next i
'        End If
'
'        For i = 1 To shifts.Count
'            If shifts(i).normalAvail > 0 Then
'                StoreNormalPiece dictNormal, shifts(i).data, shifts(i).normalAvail
'            End If
'        Next i
'    Next k
'
'    WriteDictionaryToSheet dictNormal, outputSheets("Normal")
'    WriteDictionaryToSheet dictR2, outputSheets("OT_Rule2")
'End Sub

'Public Sub ProcessRule3_WeekGT38Only()
'    If Not IsInitialized Then
'        MsgBox "Processor not properly initialized.", vbCritical
'        Exit Sub
'    End If
'
'    'PrepareAllOutputSheets
'
'    Dim dayMap As Object
'    GatherDailyBuckets dayMap
'
'    ' Convert to list for weekly processing
'    Dim normals As Collection
'    Dim dictNormal As Object: Set dictNormal = CreateObject("Scripting.Dictionary")
'    Dim k As Variant, shifts As Collection, i As Long
'    For Each k In dayMap.Keys
'        Set shifts = dayMap(k)
'        For i = 1 To shifts.Count
'            If shifts(i).rawUnits > 0 Then
'                StoreByKey dictNormal, shifts(i).data
'            End If
'        Next i
'    Next k
'
'    ExpandDictToList dictNormal, normals
'
'    ' Apply Rule 3: Week > 38 hours
'    Dim dictWeekGT38 As Object
'    Call ApplyWeeklyHoursGT38_FromRaw(dayMap, dictWeekGT38)
'
'    ' Remaining normal hours
'    Dim dictNormalFinal As Object: Set dictNormalFinal = CreateObject("Scripting.Dictionary")
'    For i = 1 To normals.Count
'        Dim pr As PayrollRowDataClass
'        Set pr = normals(i)
'        If pr.hours > 0 Then StoreByKey dictNormalFinal, pr
'    Next i
'
'    ' Write outputs
'    WriteDictionaryToSheet dictNormalFinal, outputSheets("Normal")
'    WriteDictionaryToSheet dictWeekGT38, outputSheets("OT_Rule3")
'
'    LogError "Write", "", 0, "Rule3Only", "Rows written", _
'        "Normal=" & outputSheets("Normal").UsedRange.Rows.Count - 1 & _
'        ", WeekGT38=" & outputSheets("OT_Rule3").UsedRange.Rows.Count - 1
'End Sub
'
'Public Sub ProcessRule4Only()
'    If Not IsInitialized Then
'        MsgBox "Processor not properly initialized.", vbCritical
'        Exit Sub
'    End If
'
'    PrepareAllOutputSheets
'
'    ' ------------------------
'    ' Declarations (once only)
'    ' ------------------------
'    Dim dayMap As Object
'    Dim r4Days As Object
'    Dim weekMap As Object
'    Dim k As Variant
'    Dim wk As Variant
'    Dim shifts As Collection
'    Dim emp As String
'    Dim wkKey As String
'    Dim dsk As Long
'    Dim dayStr As String
'    Dim daysDict As Object
'    Dim arr() As String
'    Dim c As Long
'    Dim ds As Variant
'    Dim i As Long, j As Long
'    Dim tmp As String
'    Dim dictNormal As Object
'    Dim dictR4 As Object
'    Dim idx As Long
'    Dim take As Double
'    Dim isR4 As Boolean
'    Dim dk As String
'    Dim inner As Object
'
'    ' 1) Gather shifts into employee-day buckets
'    GatherDailyBuckets dayMap
'
'    ' 2) Build (emp, weekSortKey) -> set of worked day keys (yyyymmdd)
'    Set r4Days = CreateObject("Scripting.Dictionary")
'    Set weekMap = CreateObject("Scripting.Dictionary")
'
'    For Each k In dayMap.Keys
'        Set shifts = dayMap(k)
'        If shifts.Count > 0 Then
'            emp = CStr(shifts(1).EmployeeCode)
'            wkKey = weekKey2(emp, CLng(shifts(1).weekSortKey))
'            dsk = CLng(shifts(1).dateSortKey)
'            dayStr = CStr(dsk)
'
'            If Not weekMap.Exists(wkKey) Then
'                Set inner = CreateObject("Scripting.Dictionary")
'                weekMap.Add wkKey, inner
'            End If
'            weekMap(wkKey)(dayStr) = True
'        End If
'    Next k
'
'    ' 3) For each (emp, weekSortKey), sort days asc and mark 6th & 7th as Rule 4 days
'    For Each wk In weekMap.Keys
'        Set daysDict = weekMap(wk)
'
'        ReDim arr(1 To daysDict.Count)
'        c = 0
'        For Each ds In daysDict.Keys
'            c = c + 1
'            arr(c) = CStr(ds)
'        Next ds
'
'        ' Sort ascending numerically by yyyymmdd
'        For i = 1 To UBound(arr) - 1
'            For j = i + 1 To UBound(arr)
'                If CLng(arr(i)) > CLng(arr(j)) Then
'                    tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
'                End If
'            Next j
'        Next i
'
'        ' Mark 6th and 7th worked days
'        For i = 6 To UBound(arr)
'            emp = Split(CStr(wk), "|")(0)      ' wkKey uses vbLf join: emp & vbLf & weekSortKey
'            dsk = CLng(arr(i))
'            dk = dayKey(emp, dsk)               ' builds emp & vbLf & yyyymmdd
'            r4Days(dk) = True
'        Next i
'    Next wk
'
'    ' 4) Allocate: for marked days => code 610; others => Normal
'    Set dictNormal = CreateObject("Scripting.Dictionary")
'    Set dictR4 = CreateObject("Scripting.Dictionary")
'
'    For Each k In dayMap.Keys
'        Set shifts = dayMap(k)
'        isR4 = r4Days.Exists(CStr(k))
'        For idx = 1 To shifts.Count
'            take = shifts(idx).normalAvail
'            If take > 0 Then
'                If isR4 Then
'                    ' All hours on the 6th/7th worked day go to code 610
'                    StoreR4Piece dictR4, shifts(idx).data, take, "610"
'                Else
'                    ' Otherwise the whole shift is Normal
'                    StoreNormalPiece dictNormal, shifts(idx).data, take
'                End If
'                shifts(idx).normalAvail = 0
'            End If
'        Next idx
'    Next k
'
'    ' 5) Write outputs
'    WriteDictionaryToSheet dictNormal, outputSheets("Normal")
'    If outputSheets.Exists("OT_Rule4") Then
'        WriteDictionaryToSheet dictR4, outputSheets("OT_Rule4")
'    Else
'        LogError "Write", "", 0, "ProcessRule4Only", "Missing output sheet key 'OT_Rule4'", _
'                 "Call RegisterOutputSheet ""OT_Rule4"", Worksheets(""OTRule4"")"
'    End If
'
'    ' Summary
'    LogError "Write", "", 0, "Rule4Only", "Rows written", _
'        "Normal=" & outputSheets("Normal").UsedRange.Rows.Count - 1 & _
'        ", R4=" & IIf(outputSheets.Exists("OT_Rule4"), outputSheets("OT_Rule4").UsedRange.Rows.Count - 1, 0)
'End Sub

' Helper used by Rule 4: store Rule 4 units with supplied code (610)
Private Sub StoreR4Piece(ByRef dict As Object, ByRef d As PayrollRowDataClass, units As Double, payrollCode As String)
    Dim part As PayrollRowDataClass
    Set part = New PayrollRowDataClass
    CopyPayrollRowData d, part
    part.hours = units
    part.payrollCode = payrollCode

    Dim why As String
    If Not ValidateForNormalOrR1(part, why) Then
        LogError "Rule4", part.EmployeeCode, part.DateIn, "StoreR4Piece", why, _
                 "Emp=" & part.EmployeeCode & ", Co=" & part.companyCode & _
                 ", PayClass=" & part.PayClassCode & ", PayCode=" & part.payrollCode
        Exit Sub
    End If

    StoreByKey dict, part
End Sub
' Rule 1: Shift > 5h (per-shift), independent of other rules
Private Sub ProcessRule1_Independent(ByVal dayMap As Object, ByRef outR1 As Object)
    Set outR1 = CreateObject("Scripting.Dictionary")

    Dim k As Variant, shifts As Collection, i As Long
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        For i = 1 To shifts.Count
            Dim r1u As Double
            r1u = Application.WorksheetFunction.Max(0, shifts(i).r1Units)  ' units = hours*10000
            If r1u > 0 Then
                ' StoreR1Piece sets payrollCode = "615"
                StoreR1Piece outR1, shifts(i).data, r1u
            End If
        Next i
    Next k
End Sub
' Rule 2: Daily total > 11.5h, independent of other rules
' Uses rawUnits only and allocates from end-of-day backwards.
' ===== REPLACEMENT =====
' Rule 2: Daily total > 11.5h, independent of other rules
' Allocates from the latest shift backward: AFTER 2h (higher band) first, then FIRST 2h.
' ===== REPLACEMENT: ProcessRule2_DailyIndependent =====
Private Sub ProcessRule2_DailyIndependent(ByVal dayMap As Object, ByRef outR2 As Object)
    Set outR2 = CreateObject("Scripting.Dictionary")

    Dim k As Variant, shifts As Collection, i As Long
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        ' Sum the day from raw units
        Dim dayUnits As Double: dayUnits = 0
        For i = 1 To shifts.Count
            dayUnits = dayUnits + shifts(i).rawUnits
        Next i

        ' Units beyond 11.5h
        Dim r2Units As Double
        r2Units = Application.WorksheetFunction.Max(0, dayUnits - 11.5 * 10000)

        ' Split first 2h / after 2h
        Dim first2 As Double: first2 = Application.WorksheetFunction.Min(20000, r2Units)
        Dim after2 As Double: after2 = Application.WorksheetFunction.Max(0, r2Units - 20000)

        ' Day-of-week codes
        Dim code1 As String, code2 As String, tmpA As Double, tmpB As Double
        Call ApplyOvertimeRule2( _
            shifts(1).data.DateTimeIn, shifts(1).data.DateTimeIn + (12# / 24#), _
            shifts(1).data.dayOfWeek, _
            code1, code2, tmpA, tmpB)

        ' Allocate AFTER 2h first, from latest shift backward
        Dim need As Double, take As Double
        need = after2
        For i = shifts.Count To 1 Step -1
            take = Application.WorksheetFunction.Min(need, shifts(i).normalAvail)
            If take > 0 Then
                StoreR2Piece outR2, shifts(i).data, take, code2
                shifts(i).normalAvail = shifts(i).normalAvail - take
                need = need - take
                If need <= 0 Then Exit For
            End If
        Next i

        ' Then FIRST 2h
        need = first2
        For i = shifts.Count To 1 Step -1
            take = Application.WorksheetFunction.Min(need, shifts(i).normalAvail)
            If take > 0 Then
                StoreR2Piece outR2, shifts(i).data, take, code1
                shifts(i).normalAvail = shifts(i).normalAvail - take
                need = need - take
                If need <= 0 Then Exit For
            End If
        Next i
    Next k
End Sub

' Rule 3: Weekly total > 38h (by employee & weekSortKey), independent of other rules
' Allocates from the latest day backwards, using the same day-of-week codes as Rule 2.
Private Sub ProcessRule3_WeeklyIndependent(ByVal dayMap As Object, ByRef outW3 As Object)
    Set outW3 = CreateObject("Scripting.Dictionary")

    ' Build (emp, weekSortKey) -> collection of shifts
    Dim weekMap As Object: Set weekMap = CreateObject("Scripting.Dictionary")
    Dim k As Variant, shifts As Collection, i As Long, wkKey As String
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        For i = 1 To shifts.Count
            wkKey = weekKey2(shifts(i).EmployeeCode, CLng(shifts(i).weekSortKey))
            If Not weekMap.Exists(wkKey) Then weekMap.Add wkKey, New Collection
            weekMap(wkKey).Add shifts(i)
        Next i
    Next k

    ' Process each week bucket
    Dim wk As Variant
    For Each wk In weekMap.keys
        Dim lst As Collection: Set lst = weekMap(wk)

        ' Total raw units in the week
        Dim total As Double: total = 0
        For i = 1 To lst.Count: total = total + lst(i).rawUnits: Next i

        Dim needUnits As Double: needUnits = Application.WorksheetFunction.Max(0, total - 38# * 10000)
        If needUnits <= 0 Then GoTo NextWeek

        ' Sort by date ascending so we can allocate from end (latest first)
        Dim arr() As Variant: ReDim arr(1 To lst.Count)
        Dim a As Long, b As Long, tmp As Variant
        For i = 1 To lst.Count: Set arr(i) = lst(i): Next i
        For a = 1 To UBound(arr) - 1
            For b = a + 1 To UBound(arr)
                If arr(a).dateSortKey > arr(b).dateSortKey Then
                    Set tmp = arr(a): Set arr(a) = arr(b): Set arr(b) = tmp
                End If
            Next b
        Next a

        ' Weekly split across the whole week
        Dim wFirst2 As Double, wAfter2 As Double
        wFirst2 = Application.WorksheetFunction.Min(20000, needUnits)
        wAfter2 = Application.WorksheetFunction.Max(0, needUnits - 20000)

        Dim need As Double, take As Double, cap As Double
        Dim code1 As String, code2 As String, tmpA As Double, tmpB As Double

        ' AFTER first 2h (code2) from the latest day backward
        need = wAfter2
        For i = UBound(arr) To 1 Step -1
            cap = arr(i).rawUnits
            take = Application.WorksheetFunction.Min(need, cap)
            If take > 0 Then
                Call ApplyOvertimeRule2(arr(i).data.DateTimeIn, arr(i).data.DateTimeIn + (12# / 24#), _
                                        arr(i).data.dayOfWeek, code1, code2, tmpA, tmpB)
                StoreR2Piece outW3, arr(i).data, take, code2
                need = need - take
                If need <= 0 Then Exit For
            End If
        Next i

        ' FIRST 2h (code1) from the latest day backward
        need = wFirst2
        For i = UBound(arr) To 1 Step -1
            cap = arr(i).rawUnits
            take = Application.WorksheetFunction.Min(need, cap)
            If take > 0 Then
                Call ApplyOvertimeRule2(arr(i).data.DateTimeIn, arr(i).data.DateTimeIn + (12# / 24#), _
                                        arr(i).data.dayOfWeek, code1, code2, tmpA, tmpB)
                StoreR2Piece outW3, arr(i).data, take, code1
                need = need - take
                If need <= 0 Then Exit For
            End If
        Next i
NextWeek:
    Next wk
End Sub

' Rule 4: Days 6 & 7 only (per employee & weekSortKey), independent of other rules
' All hours on the 6th and 7th worked calendar days get payroll code 610.
' Uses rawUnits only (does NOT consume/alter other pools).
Private Sub ProcessRule4_DaysGT5Independent(ByVal dayMap As Object, ByRef outR4 As Object)
    Set outR4 = CreateObject("Scripting.Dictionary")

    Dim weekMap As Object: Set weekMap = CreateObject("Scripting.Dictionary")
    Dim k As Variant, shifts As Collection, i As Long
    Dim emp As String, wkKey As String, dayStr As String

    ' Build (emp, weekSortKey) -> (yyyymmdd -> Collection of shifts on that day)
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        If shifts.Count > 0 Then
            emp = CStr(shifts(1).EmployeeCode)
            wkKey = weekKey2(emp, CLng(shifts(1).weekSortKey))

            If Not weekMap.Exists(wkKey) Then
                weekMap.Add wkKey, CreateObject("Scripting.Dictionary")
            End If

            For i = 1 To shifts.Count
                dayStr = CStr(CLng(shifts(i).dateSortKey)) ' yyyymmdd
                If Not weekMap(wkKey).Exists(dayStr) Then
                    Set weekMap(wkKey)(dayStr) = New Collection
                End If
                weekMap(wkKey)(dayStr).Add shifts(i)
            Next i
        End If
    Next k

    ' For each week bucket, sort days ascending and convert ONLY day #6 and #7
    Dim wk As Variant, daysDict As Object
    Dim arr() As String, c As Long, ds As Variant, a As Long, b As Long, tmp As String
    Dim col As Collection, take As Double, j As Long
    Dim maxIdx As Long

    For Each wk In weekMap.keys
        Set daysDict = weekMap(wk)

        ' Collect and sort calendar days ascending (yyyymmdd)
        ReDim arr(1 To daysDict.Count)
        c = 0
        For Each ds In daysDict.keys
            c = c + 1
            arr(c) = CStr(ds)
        Next ds

        For a = 1 To UBound(arr) - 1
            For b = a + 1 To UBound(arr)
                If CLng(arr(a)) > CLng(arr(b)) Then
                    tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
                End If
            Next b
        Next a

        ' Clamp to exactly day #6 and #7 (ignore 8th+ if present)
        If UBound(arr) >= 7 Then
            maxIdx = 7
        Else
            maxIdx = UBound(arr)
        End If

        If maxIdx >= 6 Then
            For a = 6 To maxIdx
                Set col = daysDict(arr(a))
                For j = 1 To col.Count
                    ' Independent allocation: use rawUnits (do not touch normalAvail)
                    take = col(j).rawUnits
                    If take > 0 Then
                        ' 610 is fixed for Rule 4 (6th/7th day only)
                        StoreR4Piece outR4, col(j).data, take, "610"
                    End If
                Next j
            Next a
        End If
    Next wk
End Sub
Public Sub ProcessAllowances()
    If LookupSheet Is Nothing Or HolidaysSheet Is Nothing Or InputSheet Is Nothing Then
        MsgBox "Allowances: required sheets not initialized (DataIn/Lookup/Holidays).", vbCritical
        Exit Sub
    End If
    If outputSheets Is Nothing Or Not outputSheets.Exists("Allowances") Then
        MsgBox "Allowances: output sheet key 'Allowances' is not registered.", vbCritical
        Exit Sub
    End If

    Dim wsIn As Worksheet, wsLookup As Worksheet, wsHol As Worksheet, wsOut As Worksheet
    Set wsIn = InputSheet
    Set wsLookup = LookupSheet
    Set wsHol = HolidaysSheet
    Set wsOut = outputSheets("Allowances")

    ' Prepare output
    PrepareAllowancesSheet_ wsOut

    ' Aggregation dictionary: key -> units
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    ' Helper (reuse existing LookupManager for company and holiday check)
    Dim lm As LookupManager: Set lm = New LookupManager
    Set lm.LookupSheet = wsLookup
    Set lm.ADPSheet = ADPSheet
    Set lm.HolidaysSheet = wsHol

    Dim lastRow As Long, i As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then GoTo WriteOut

    For i = 2 To lastRow
        ' ---- Input row fields ----
        Dim ownership As Variant, payrollExportCode As String, weekEnding As Variant
        Dim empCode As String, posCode As Variant, glNumber As Variant
        Dim vDateIn As Variant, vDateOut As Variant, vTimeIn As Variant, vTimeOut As Variant
        Dim payRate As Variant

        ownership = wsIn.Cells(i, 1).value
        payrollExportCode = CStr(wsIn.Cells(i, 2).value)
        weekEnding = wsIn.Cells(i, 3).value
        empCode = CStr(wsIn.Cells(i, 4).value)
        glNumber = wsIn.Cells(i, 6).value
        vDateIn = wsIn.Cells(i, 7).value
        vDateOut = wsIn.Cells(i, 8).value
        vTimeIn = wsIn.Cells(i, 9).value
        vTimeOut = wsIn.Cells(i, 10).value
        payRate = wsIn.Cells(i, 11).value  ' not used by allowances, kept for parity

        ' ---- Derive dates/times ----
        Dim dIn As Date, dOut As Date, dtIn As Date, dtOut As Date
        dIn = ParseYYMMDDandYYYYMMDDtoDate(vDateIn)
        dOut = ParseYYMMDDandYYYYMMDDtoDate(vDateOut)

        If IsDate(vTimeIn) Then
            dtIn = dIn + CDbl(vTimeIn)
        ElseIf IsNumeric(vTimeIn) And vTimeIn >= 0 And vTimeIn < 1 Then
            dtIn = dIn + CDbl(vTimeIn)
        Else
            dtIn = dIn
        End If

        If IsDate(vTimeOut) Then
            dtOut = dOut + CDbl(vTimeOut)
        ElseIf IsNumeric(vTimeOut) And vTimeOut >= 0 And vTimeOut < 1 Then
            dtOut = dOut + CDbl(vTimeOut)
        Else
            dtOut = dOut
        End If

        ' overnight safety
        If dtOut < dtIn Then dtOut = dtOut + 1

        ' ---- Basic derived fields ----
        Dim companyCode As String
        companyCode = lm.GetCompanyCode(ownership)

        Dim dayOfWeek As Integer
        dayOfWeek = Weekday(dIn, vbMonday) ' Mon=1  Sun=7

        Dim entryDate As String, fromDate As String, toDate As String
        entryDate = ParseYYMMDDtoDDMMYY(weekEnding)
        fromDate = Format(dIn, "DDMMYY")
        toDate = Format(dOut, "DDMMYY")

        Dim weekSortKey As Long, dateSortKey As Long
        weekSortKey = ParseYYMMDDToYYYYMMDD(weekEnding)
        dateSortKey = CLng(Format(dIn, "yyyymmdd"))

        ' ---- Cost Centre (suffix + PayrollExportCode) ----
        ' Lookup range name: CostCodeSuffix (GL -> suffix)
        Dim suffix As Variant, costCentre As String
        On Error Resume Next
        suffix = Application.VLookup(CLng(glNumber), wsLookup.Range("CostCodeSuffix"), 2, False)
        On Error GoTo 0
        If isError(suffix) Then suffix = ""
        costCentre = CStr(suffix) & payrollExportCode

        ' ---- Holiday flag (for A101 logic) ----
        Dim isHoliday As Boolean
        isHoliday = Not isError(Application.VLookup(payrollExportCode & Format(dIn, "YYMMDD"), wsHol.Range("A:B"), 2, False))

        ' ---- Compute allowances (whole hours * 100) ----
        Dim allowUnits As Long, allowCode As String

        ' Reset per row
        allowUnits = 0
        allowCode = ""

        ' A101 (MonFri, not holiday, hours before 06:00)
        If dayOfWeek >= 1 And dayOfWeek <= 5 And Not isHoliday Then
            Dim t As Date, wholeHrs As Long
            t = dtIn
            wholeHrs = 0
            Do While t < dtOut
                If Hour(t) < 6 Then wholeHrs = wholeHrs + 1
                t = t + TimeSerial(1, 0, 0)
            Loop
            If wholeHrs > 0 Then
                allowCode = "A101"
                allowUnits = allowUnits + wholeHrs * 100   ' original scaling
                ' Aggregate A101
                If allowUnits > 0 Then
                    Dim keyA101 As String
                    keyA101 = companyCode & vbLf & empCode & vbLf & "A" & vbLf & entryDate & vbLf & _
                              "A101" & vbLf & costCentre & vbLf & "" & vbLf & "" & vbLf & _
                              fromDate & vbLf & toDate & vbLf & CStr(weekSortKey) & vbLf & CStr(dateSortKey)
                    If dict.Exists(keyA101) Then
                        dict(keyA101) = CLng(dict(keyA101)) + allowUnits
                    Else
                        dict.Add keyA101, allowUnits
                    End If
                End If
            End If
        End If

        ' A100 (time after 22:00)
        ' NOTE: Your original code did NOT exclude holidays here (left the check commented out),
        ' so we keep that behaviour.
        Dim varDateTime As Date, wholeHrs2 As Long
        varDateTime = dIn + TimeSerial(22, 0, 0)
        wholeHrs2 = 0
        If dtOut >= varDateTime Then
            Do While varDateTime < dtOut
                wholeHrs2 = wholeHrs2 + 1
                varDateTime = varDateTime + TimeSerial(1, 0, 0)
            Loop
            If wholeHrs2 > 0 Then
                Dim unitsA100 As Long: unitsA100 = wholeHrs2 * 100
                Dim keyA100 As String
                keyA100 = companyCode & vbLf & empCode & vbLf & "A" & vbLf & entryDate & vbLf & _
                          "A100" & vbLf & costCentre & vbLf & "" & vbLf & "" & vbLf & _
                          fromDate & vbLf & toDate & vbLf & CStr(weekSortKey) & vbLf & CStr(dateSortKey)
                If dict.Exists(keyA100) Then
                    dict(keyA100) = CLng(dict(keyA100)) + unitsA100
                Else
                    dict.Add keyA100, unitsA100
                End If
            End If
        End If
    Next i

WriteOut:
    ' Write aggregated allowances (filter out Cost Centres starting with "M")
    WriteAllowancesDictToSheet_ dict, wsOut
End Sub
' Prepare AllowancesOut headers and formats
Private Sub PrepareAllowancesSheet_(ByVal ws As Worksheet)
    With ws
        .Cells.Clear
        .Cells(1, 1).value = "Company Code"
        .Cells(1, 2).value = "Employee Code"
        .Cells(1, 3).value = "Record Type"
        .Cells(1, 4).value = "Entry Date"
        .Cells(1, 5).value = "Allowance Code"
        .Cells(1, 6).value = "Amount/Units"
        .Cells(1, 7).value = "Cost Centre"
        .Cells(1, 8).value = "Notation 1"
        .Cells(1, 9).value = "Notation 2"
        .Cells(1, 10).value = "From Date"
        .Cells(1, 11).value = "To Date"
        .Cells(1, 12).value = "Week Sort Key"
        .Cells(1, 13).value = "Date Sort Key"

        .Columns(2).NumberFormat = "@"
        .Columns(4).NumberFormat = "@"
        .Columns(5).NumberFormat = "@"
        .Columns(6).NumberFormat = "@"
        .Columns(10).NumberFormat = "@"
        .Columns(11).NumberFormat = "@"
        .Columns(12).NumberFormat = "0"
        .Columns(13).NumberFormat = "0"
    End With
End Sub

' Write allowances aggregation to the output, skipping Cost Centres starting with "M"
Private Sub WriteAllowancesDictToSheet_(ByVal dict As Object, ByVal ws As Worksheet)
    Dim r As Long: r = 2
    Dim k As Variant, parts() As String
    For Each k In dict.keys
        parts = Split(CStr(k), vbLf)
        ' parts(0..11) must exist
        If UBound(parts) >= 11 Then
            ' Filter: cost centre not starting with "M"
            If Len(parts(5)) > 0 And UCase$(Left$(Trim$(parts(5)), 1)) <> "M" Then
                ws.Cells(r, 1).value = parts(0)                ' Company Code
                ws.Cells(r, 2).value = parts(1)                ' Employee Code
                ws.Cells(r, 3).value = parts(2)                ' Record Type ("A")
                ws.Cells(r, 4).value = parts(3)                ' Entry Date (DDMMYY)
                ws.Cells(r, 5).value = parts(4)                ' Allowance Code (A100/A101)
                ws.Cells(r, 6).value = dict(k)                 ' Amount/Units (hours*100)
                ws.Cells(r, 7).value = parts(5)                ' Cost Centre
                ws.Cells(r, 8).value = parts(6)                ' Notation 1
                ws.Cells(r, 9).value = parts(7)                ' Notation 2
                ws.Cells(r, 10).value = parts(8)               ' From Date
                ws.Cells(r, 11).value = parts(9)               ' To Date
                ws.Cells(r, 12).value = parts(10)              ' Week Sort Key
                ws.Cells(r, 13).value = parts(11)              ' Date Sort Key
                r = r + 1
            End If
        End If
    Next
    ws.Columns.AutoFit
End Sub
' Appends aggregated allowances to the output (no clearing).
' Keeps the same column order and the "skip Cost Centres starting with M" rule.
Private Sub WriteAllowancesDictAppend_(ByVal dict As Object, ByVal ws As Worksheet)
    Dim r As Long: r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If r < 2 Then r = 2

    Dim k As Variant, parts() As String
    For Each k In dict.keys
        parts = Split(CStr(k), vbLf)
        If UBound(parts) >= 11 Then
            ' Filter: cost centre not starting with "M"
            If Len(parts(5)) > 0 And UCase$(Left$(Trim$(parts(5)), 1)) <> "M" Then
                ws.Cells(r, 1).value = parts(0)   ' Company Code
                ws.Cells(r, 2).value = parts(1)   ' Employee Code
                ws.Cells(r, 3).value = parts(2)   ' Record Type ("A")
                ws.Cells(r, 4).value = parts(3)   ' Entry Date (DDMMYY)
                ws.Cells(r, 5).value = parts(4)   ' Allowance Code (A102)
                ws.Cells(r, 6).value = dict(k)    ' Amount/Units (money  100)
                ws.Cells(r, 7).value = parts(5)   ' Cost Centre
                ws.Cells(r, 8).value = parts(6)   ' Notation 1 (hours text)
                ws.Cells(r, 9).value = parts(7)   ' Notation 2
                ws.Cells(r, 10).value = parts(8)  ' From Date
                ws.Cells(r, 11).value = parts(9)  ' To Date
                ws.Cells(r, 12).value = parts(10) ' Week Sort Key
                ws.Cells(r, 13).value = parts(11) ' Date Sort Key
                r = r + 1
            End If
        End If
    Next
    ws.Columns.AutoFit
End Sub
' ===== PATCH: Replace entire ProcessRule1_AsAllowance =====
Private Sub ProcessRule1_AsAllowance(ByVal dayMap As Object)
    Dim wsOut As Worksheet
    Set wsOut = outputSheets("Allowances")
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim k As Variant, shifts As Collection, i As Long
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        For i = 1 To shifts.Count
            Dim r1Units As Double
            r1Units = Application.WorksheetFunction.Max(0, shifts(i).r1Units) ' units = hours * 10000
            If r1Units > 0 Then
                Dim d As PayrollRowDataClass: Set d = shifts(i).data
                Dim r1Hours As Double: r1Hours = Round(r1Units / 10000#, 2)
                Dim basePay As Double: basePay = ValidateNumeric(d.payRate, 0)
                Dim amountUnits As Long: amountUnits = CLng(Round(basePay * r1Hours * 100#, 0))
                Dim note1 As String
                note1 = "Number of hours = " & Format$(r1Hours, "0.00") & " @ " & Format$(basePay, "0.00") & " per hour"

                ' Suppression key check
                Dim isError As Boolean: isError = IsErrorShift(d)

                If isError Then
                    ' Log suppression to Errors with Hours = r1Hours
                    Call LogErrorWithHours("Allowance", d.EmployeeCode, d.DateIn, r1Hours, "AllowanceRemoved", _
                                           "Suppressed A102 (error shift)", _
                                           "Site=" & d.payrollExportCode & "; Co=" & d.companyCode & "; CostCentre=" & d.costCentre)
                Else
                    ' Aggregate A102
                    Dim key As String
                    key = d.companyCode & vbLf & d.EmployeeCode & vbLf & "A" & vbLf & d.entryDate & vbLf & _
                          "A102" & vbLf & d.costCentre & vbLf & note1 & vbLf & "" & vbLf & _
                          d.fromDate & vbLf & d.toDate & vbLf & CStr(d.weekSortKey) & vbLf & CStr(d.dateSortKey)
                    If dict.Exists(key) Then
                        dict(key) = CLng(dict(key)) + amountUnits
                    Else
                        dict.Add key, amountUnits
                    End If
                End If
            End If
        Next i
    Next k

    ' Append (do not clear) to AllowancesOut
    WriteAllowancesDictAppend_ dict, wsOut
End Sub

Private Sub ProcessExportDeduped(ByVal dayMap As Object)
    Dim outOT As Object, outNormal As Object
    Set outOT = CreateObject("Scripting.Dictionary")
    Set outNormal = CreateObject("Scripting.Dictionary")

    Dim k As Variant, wk As Variant, dayKey As Variant
    Dim shifts As Collection, col As Collection
    Dim i As Long, j As Long

    ' --- (PH) Public Holiday -> 603 (site-aware; consumes minutes first)
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        If Not shifts Is Nothing Then
            For i = 1 To shifts.Count
                If IsPublicHolidayShift(shifts(i)) Then
                    If shifts(i).normalAvail > 0 Then
                        ' Diagnostics
                        Debug.Print "PH-CONSUME", shifts(i).EmployeeCode, _
                                    "Date:", Format(shifts(i).data.DateIn, "dd/mm/yyyy"), _
                                    "Units:", CLng(shifts(i).normalAvail), _
                                    "Site:", shifts(i).payrollExportCode
                        StoreWithCode outOT, outNormal, shifts(i).data, shifts(i).normalAvail, "603"
                        shifts(i).normalAvail = 0
                    End If
                End If
            Next i
        End If
    Next k

    
    ' Build weekMap: Collection of ShiftPieces
    Dim weekMap As Object: Set weekMap = CreateObject("Scripting.Dictionary")
    Dim emp As String, wkKey As String
    For Each dayKey In dayMap.keys
        Set shifts = dayMap(dayKey)
        If Not shifts Is Nothing Then
            If shifts.Count > 0 Then
                emp = CStr(shifts(1).EmployeeCode)
                wkKey = weekKey2(emp, CLng(shifts(1).weekSortKey))
                If Not weekMap.Exists(wkKey) Then weekMap.Add wkKey, CreateObject("Scripting.Dictionary")
                If Not weekMap(wkKey).Exists(dayKey) Then Set weekMap(wkKey)(dayKey) = New Collection
                For j = 1 To shifts.Count
                    weekMap(wkKey)(dayKey).Add shifts(j)
                Next j
            End If
        End If
    Next dayKey

    ' For each employee/week, sort days ASC and consume ONLY the 6th/7th days WITH remaining minutes
    '(R4) 6th/7th worked day -> 610 (per employee/week), counting ONLY days with remaining minutes
    Dim daysDict As Object
    Dim arr() As Variant ' dynamic 2D: arr(1..3, 1..c) = [yyyymmdd; dayKey; remainingDayUnits]
    Dim c As Long, ds As Variant, a As Long, b As Long, idx As Long
    Dim dWork As Date, remainingDayUnits As Long
    Dim t1 As Variant, t2 As Variant, t3 As Variant
    
    For Each wk In weekMap.keys
        Set daysDict = weekMap(wk)

        ' Build arr dynamically on the LAST dimension (legal for ReDim Preserve)
        c = 0
        Erase arr ' ensure arr is unallocated before reuse

        
        For Each ds In daysDict.keys
            Set col = daysDict(ds)
            remainingDayUnits = 0
            If Not col Is Nothing Then
                For j = 1 To col.Count
                    remainingDayUnits = remainingDayUnits + CLng(col(j).normalAvail)
                Next j
            End If

            If remainingDayUnits > 0 Then
                c = c + 1
                If c = 1 Then
                    ReDim arr(1 To 3, 1 To 1)
                Else
                    ReDim Preserve arr(1 To 3, 1 To c)
                End If
                dWork = CDate(col(1).data.DateIn)
                arr(1, c) = CLng(Format(dWork, "yyyymmdd")) ' date sort key
                arr(2, c) = ds                              ' original dayKey
                arr(3, c) = remainingDayUnits               ' remaining units for that day
            End If
        Next ds

        ' Nothing to sort/consume if <= 1 remaining day
        If c <= 1 Then GoTo NextWeekR4

        ' Sort ascending by date (field 1), swap all three fields together
        For a = 1 To c - 1
            For b = a + 1 To c
                If CLng(arr(1, a)) > CLng(arr(1, b)) Then
                    t1 = arr(1, a): t2 = arr(2, a): t3 = arr(3, a)
                    arr(1, a) = arr(1, b): arr(2, a) = arr(2, b): arr(3, a) = arr(3, b)
                    arr(1, b) = t1:       arr(2, b) = t2:       arr(3, b) = t3
                End If
            Next b
        Next a

        ' Consume day #6 and (if exists) #7
        If c >= 6 Then
            For idx = 6 To IIf(c >= 7, 7, c)
                ds = arr(2, idx) ' original dictionary dayKey
                Set col = daysDict(ds)
                If Not col Is Nothing Then
                    For j = 1 To col.Count
                        If col(j).normalAvail > 0 Then
                            ' Diagnostics
                            Debug.Print "R4-CONSUME", col(j).EmployeeCode, _
                                        "Date:", Format(col(j).data.DateIn, "dd/mm/yyyy"), _
                                        "Units:", CLng(col(j).normalAvail), _
                                        "WeekKey:", wk
                            StoreWithCode outOT, outNormal, col(j).data, col(j).normalAvail, "610"
                            col(j).normalAvail = 0
                        End If
                    Next j
                End If
            Next idx
        End If

NextWeekR4:
    Next wk
    
    ' ==== Highest-Rated bands on remaining minutes ====
    Const DAILY_THRESHOLD_UNITS As Long = 115000 ' 11.5 h * 10000
    Const WEEKLY_THRESHOLD_UNITS As Long = 380000 ' 38 h * 10000

    ' 1) Daily bands per day
    Dim dailyA2 As Object, dailyF2 As Object
    Set dailyA2 = CreateObject("Scripting.Dictionary")
    Set dailyF2 = CreateObject("Scripting.Dictionary")

    Dim excess As Long, needF2 As Long, needA2 As Long
    Dim takeA2 As Object, takeF2 As Object, kk As Variant
    Dim dayTotal As Long

    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        dayTotal = DayMinutes(shifts)
        If dayTotal <= 0 Then GoTo NextDay

        excess = dayTotal - DAILY_THRESHOLD_UNITS
        If excess < 0 Then excess = 0
        needF2 = IIf(excess > 20000, 20000, excess)
        needA2 = IIf(excess > 20000, excess - 20000, 0)

        ' Diagnostics
        Debug.Print "DailyCheck", shifts(1).EmployeeCode, _
                    "Date:", Format(shifts(1).data.DateIn, "dd/mm/yyyy"), _
                    "DayTotal:", dayTotal, _
                    "Excess:", (dayTotal - DAILY_THRESHOLD_UNITS), _
                    "F2:", needF2, "A2:", needA2

        Set takeA2 = AllocateFromLatestInDay(shifts, needA2) ' compute only (no subtraction)
        Set takeF2 = AllocateFromLatestInDay(shifts, needF2)

        For Each kk In takeA2.keys
            dailyA2(kk) = CLng(dailyA2.item(kk)) + CLng(takeA2(kk))
        Next kk
        For Each kk In takeF2.keys
            dailyF2(kk) = CLng(dailyF2.item(kk)) + CLng(takeF2(kk))
        Next kk
NextDay:
    Next k

    ' 2) Weekly bands per employee/week (on remaining minutes)
    Dim weeklyA2 As Object, weeklyF2 As Object
    Set weeklyA2 = CreateObject("Scripting.Dictionary")
    Set weeklyF2 = CreateObject("Scripting.Dictionary")

    Dim wkTotal As Long, wExcess As Long, wNeedF2 As Long, wNeedA2 As Long
    Dim takeA2W As Object, takeF2W As Object, kk2 As Variant

    For Each wk In weekMap.keys
        Set daysDict = weekMap(wk)
        wkTotal = 0
        For Each dayKey In daysDict.keys
            wkTotal = wkTotal + DayMinutes(daysDict(dayKey))
        Next dayKey

        wExcess = wkTotal - WEEKLY_THRESHOLD_UNITS
        If wExcess < 0 Then wExcess = 0
        wNeedF2 = IIf(wExcess > 20000, 20000, wExcess)
        wNeedA2 = IIf(wExcess > 20000, wExcess - 20000, 0)

        ' Diagnostics
        Debug.Print "WeeklyCheck", wk, "Total:", wkTotal, "Excess:", wExcess, "F2:", wNeedF2, "A2:", wNeedA2
        
        Set takeA2W = AllocateFromLatestInWeekBucket(daysDict, wNeedA2) ' compute only
        Set takeF2W = AllocateFromLatestInWeekBucket(daysDict, wNeedF2)

        For Each kk2 In takeA2W.keys
            weeklyA2(kk2) = CLng(weeklyA2.item(kk2)) + CLng(takeA2W(kk2))
        Next kk2
        For Each kk2 In takeF2W.keys
            weeklyF2(kk2) = CLng(weeklyF2.item(kk2)) + CLng(takeF2W(kk2))
        Next kk2
    Next wk
    
    ' 3) Combine & emit � GLOBAL latest-first (this is where subtraction & writing happens)
    Dim bandWins As Object: Set bandWins = CreateObject("Scripting.Dictionary")
    Call CombineAndEmit_GlobalLatestFirst( _
        dayMap, weekMap, _
        dailyF2, dailyA2, weeklyF2, weeklyA2, _
        outOT, outNormal, _
        bandWins)

    DumpDictionary dayMap, "dayMap", 50
    DumpDictionary weekMap, "weekMap", 50
    DumpDictionary dailyF2, "dailyF2", 50
    DumpDictionary dailyA2, "dailyA2", 50
    DumpDictionary weeklyF2, "weeklyF2", 50
    DumpDictionary weeklyA2, "weeklyA2", 50

    ' Remaining safety (should be none), spill any leftovers as base
    For Each k In dayMap.keys
        Set shifts = dayMap(k)
        If Not shifts Is Nothing Then
            For i = 1 To shifts.Count
                If shifts(i).normalAvail > 0 Then
                    StoreWithCode outOT, outNormal, shifts(i).data, shifts(i).normalAvail, MapCodeForBand(CDate(shifts(i).data.DateIn), 0)
                    shifts(i).normalAvail = 0
                End If
            Next i
        End If
    Next k

    ' --- Write results
    WriteDictionaryToSheet outNormal, outputSheets("Normal")
    WriteDictionaryToSheet outOT, outputSheets("OT_Deduped")

    ' --- Write Band Wins Audit
    Call WriteBandWinsAudit(bandWins)
End Sub
' ========= Helpers: Highest-Rated export (per-minute winner) =========


Private Function MapCodeForBand(ByVal dt As Date, ByVal band As Integer) As String
' band: 0=base, 1=first2, 2=after2
' Mon�Fri: 600/605/606, Sat: 601/607/608, Sun: 602/602/602

    Dim dow As Integer
    dow = Weekday(dt, vbMonday) ' Mon=1..Sun=7

    Select Case dow
        Case 1, 2, 3, 4, 5 ' Mon-Fri
            If band = 2 Then
                MapCodeForBand = "606"
            ElseIf band = 1 Then
                MapCodeForBand = "605"
            Else
                MapCodeForBand = "600"
            End If
        Case 6 ' Sat
            If band = 2 Then
                MapCodeForBand = "608"
            ElseIf band = 1 Then
                MapCodeForBand = "607"
            Else
                MapCodeForBand = "601"
            End If
        Case Else ' Sun
            MapCodeForBand = "602" ' Always 602 on Sunday unless PH/610 overrides
    End Select
End Function

Private Function StoreWithCode( _
    ByRef outOT As Object, ByRef outNormal As Object, _
    ByRef d As PayrollRowDataClass, ByVal units As Double, ByVal code As String) _
    As Boolean

    StoreWithCode = False
    If units <= 0# Then Exit Function

    Dim part As PayrollRowDataClass, why As String
    Set part = New PayrollRowDataClass

    ' Copy base row + override hours/code
    CopyPayrollRowData d, part
    part.hours = units
    part.payrollCode = code

    ' Decide destination (Normal vs OT) from code family
    Dim isBase As Boolean
    isBase = IsBaseCode(code)

    ' Validate row (same rules used by Normal/R1)
    If Not ValidateForNormalOrR1(part, why) Then
        ' Mark the shift for allowance suppression and log with Hours
        MarkErrorShift part, units, why
        Exit Function
    End If

    ' Persist
    If isBase Then
        StoreByKey outNormal, part
    Else
        StoreByKey outOT, part
    End If

    StoreWithCode = True
End Function

' Sum remaining minutes (normalAvail) for a day collection
Private Function DayMinutes(ByVal shifts As Collection) As Long
    Dim i As Long, sp As ShiftPieceClass, tot As Long
    For i = 1 To shifts.Count
        Set sp = shifts(i)
        tot = tot + CLng(sp.normalAvail)
    Next i
    DayMinutes = tot
End Function

' Sort a date Collection ascending (in-place)
Private Sub SortDatesCollection(ByRef c As Collection)
    Dim arr() As Date, i As Long, j As Long, t As Date
    If c.Count <= 1 Then Exit Sub
    ReDim arr(1 To c.Count)
    For i = 1 To c.Count: arr(i) = CDate(c(i)): Next
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then t = arr(i): arr(i) = arr(j): arr(j) = t
        Next j
    Next i
    Dim tmp As New Collection
    For i = 1 To UBound(arr): tmp.Add arr(i): Next
    Set c = tmp
End Sub


' ===== REPLACEMENT: AllocateFromLatestInDay =====
' Allocate tail minutes across a day's shifts (latest minutes first).
' Returns minutesByPtr(ObjPtr(sp)) = minutes taken.
' NOTE: This version DOES NOT subtract from sp.normalAvail.
Private Function AllocateFromLatestInDay(ByVal shifts As Collection, ByVal need As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If need <= 0 Or shifts Is Nothing Then
        Set AllocateFromLatestInDay = dict
        Exit Function
    End If

    Dim i As Long, sp As ShiftPieceClass, take As Long, key As String, avail As Long
    For i = shifts.Count To 1 Step -1
        If need <= 0 Then Exit For
        Set sp = shifts(i)
        avail = CLng(sp.normalAvail)          ' <-- read-only; do NOT subtract here
        If avail <= 0 Then GoTo NextI

        take = CLng(Application.WorksheetFunction.Min(avail, need))
        If take > 0 Then
            key = CStr(ObjPtr(sp))
            If Not dict.Exists(key) Then dict.Add key, 0&
            dict(key) = dict(key) + take
            need = need - take                 ' <-- only reduce the local need
        End If
NextI:
    Next i

    Set AllocateFromLatestInDay = dict
End Function
' Allocate tail minutes across the whole week (dates descending, then latest within day)
' Returns minutesByPtr(ObjPtr(sp)) = minutes taken
Private Function AllocateFromLatestInWeek(ByVal dayMap As Object, ByVal need As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If need <= 0 Then Set AllocateFromLatestInWeek = dict: Exit Function

    ' Build a sortable array of (sortKey=yyyymmdd as Long, dayKey=original dictionary key)
    Dim keys As Variant, k As Variant, shifts As Collection
    Dim arr() As Variant, c As Long

    keys = dayMap.keys
    If IsEmpty(keys) Then Set AllocateFromLatestInWeek = dict: Exit Function

    ' Pre-size to number of keys; we may shrink after we populate
    ReDim arr(1 To UBound(keys) - LBound(keys) + 1, 1 To 2)
    c = 0

    For Each k In keys
        Set shifts = dayMap(k)
        If Not shifts Is Nothing Then
            If shifts.Count > 0 Then
                ' Use the actual date from the first piece in the day
                c = c + 1
                arr(c, 1) = CLng(Format(CDate(shifts(1).data.DateIn), "yyyymmdd"))
                arr(c, 2) = k  ' keep original key for lookup
            End If
        End If
    Next k

    If c = 0 Then Set AllocateFromLatestInWeek = dict: Exit Function

    ' Shrink array to the number of populated rows
    If c < UBound(arr, 1) Then
        ReDim Preserve arr(1 To c, 1 To 2)
    End If

    ' Sort by date descending (latest first)
    Dim i As Long, j As Long, tmp1 As Variant, tmp2 As Variant
    For i = 1 To c - 1
        For j = i + 1 To c
            If CLng(arr(j, 1)) > CLng(arr(i, 1)) Then
                tmp1 = arr(i, 1): arr(i, 1) = arr(j, 1): arr(j, 1) = tmp1
                tmp2 = arr(i, 2): arr(i, 2) = arr(j, 2): arr(j, 2) = tmp2
            End If
        Next j
    Next i

    ' Allocate minutes from latest day first; within the day, latest piece first (reverse iteration)
    Dim dayKey As Variant, take As Long, sp As ShiftPieceClass, key As String
    Dim idx As Long, m As Long

    For idx = 1 To c
        If need <= 0 Then Exit For

        dayKey = arr(idx, 2)
        Set shifts = dayMap(dayKey)
        If shifts Is Nothing Then GoTo NextDay

        For m = shifts.Count To 1 Step -1
            If need <= 0 Then Exit For
            Set sp = shifts(m)

            If sp.normalAvail > 0 Then
                take = CLng(Application.WorksheetFunction.Min(sp.normalAvail, need))
                If take > 0 Then
                    key = CStr(ObjPtr(sp))
                    If Not dict.Exists(key) Then dict.Add key, 0&
                    dict(key) = dict(key) + take
                    need = need - take
                End If
            End If
        Next m
NextDay:
    Next idx

    Set AllocateFromLatestInWeek = dict
End Function

' Public Holiday check (site-aware, key = payrollExportCode & YYMMDD)
' ===== REPLACEMENT: IsPublicHolidayShift =====
' Site-aware PH check + explicit logging for diagnostics.
Private Function IsPublicHolidayShift(ByRef sp As ShiftPieceClass) As Boolean
    Dim key As String, v As Variant
    If HolidaysSheet Is Nothing Then Exit Function

    ' Key: payrollExportCode (as text) & YYMMDD
    key = CStr(sp.payrollExportCode) & Format(sp.DateIn, "YYMMDD")

    On Error Resume Next
    v = Application.VLookup(key, HolidaysSheet.Range("A:B"), 2, False)
    On Error GoTo 0

    IsPublicHolidayShift = Not isError(v)

    ' Diagnostics: print PH hits
    If IsPublicHolidayShift Then
        Debug.Print "PH-HIT", sp.EmployeeCode, _
                    "Date:", Format(sp.data.DateIn, "dd/mm/yyyy"), _
                    "Site:", sp.payrollExportCode, _
                    "Key:", key
    End If
End Function

' Create or get a worksheet by name (used for audit)
Private Function GetOrCreateSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = name
    End If
    Set GetOrCreateSheet = ws
End Function
' ========= End helpers =========

' ===== Helpers for global combiner =====
Private Function SumSelected(ByVal dict As Object, ByVal allowed As Object) As Long
    Dim k As Variant, s As Long
    s = 0
    If Not dict Is Nothing Then
        For Each k In dict.keys
            If allowed.Exists(CStr(k)) Then s = s + CLng(dict(k))
        Next k
    End If
    SumSelected = s
End Function
' Safe sorter: returns a new Collection (no in-place assignment)
Private Function SortedPiecesLatestFirst(ByVal pcs As Collection) As Collection
    Dim i As Long, j As Long
    Dim outC As New Collection
    If pcs Is Nothing Or pcs.Count = 0 Then
        Set SortedPiecesLatestFirst = outC
        Exit Function
    End If

    ' Copy to array of ShiftPieceClass refs
    Dim arr() As ShiftPieceClass
    ReDim arr(1 To pcs.Count)
    For i = 1 To pcs.Count
        Set arr(i) = pcs(i)
    Next i

    ' Sort descending by dateSortKey, then by DateTimeOut
    Dim tmp As ShiftPieceClass
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i).dateSortKey < arr(j).dateSortKey Or _
               (arr(i).dateSortKey = arr(j).dateSortKey And arr(i).DateTimeOut < arr(j).DateTimeOut) Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    ' Build result collection in sorted order
    For i = 1 To UBound(arr)
        outC.Add arr(i)
    Next i
    Set SortedPiecesLatestFirst = outC
End Function

Private Sub CombineAndEmit_GlobalLatestFirst( _
    ByRef dayMap As Object, _
    ByRef weekMap As Object, _
    ByRef dailyF2 As Object, ByRef dailyA2 As Object, _
    ByRef weeklyF2 As Object, ByRef weeklyA2 As Object, _
    ByRef outOT As Object, ByRef outNormal As Object, _
    ByRef bandWins As Object)

    Dim wk As Variant, daysDict As Object
    Dim dayKey As Variant, col As Collection
    Dim pieces As Collection
    Dim i As Long

    Dim f2Taken As Object, a2Taken As Object, baseTaken As Object
    Dim ptr As String
    Dim workDate As Date
    Dim wantA2 As Long, wantF2 As Long, take As Long
    Dim dow As Integer
    Dim ok As Boolean

    ' helper dict getters already exist in class: DictGetLong

    For Each wk In weekMap.keys
        Set daysDict = weekMap(wk)

        ' Build flat list of remaining pieces (normalAvail > 0) in this employee/week
        Set pieces = New Collection
        For Each dayKey In daysDict.keys
            Set col = daysDict(dayKey)
            If Not col Is Nothing Then
                For i = 1 To col.Count
                    If col(i).normalAvail > 0 Then pieces.Add col(i)
                Next i
            End If
        Next dayKey

        If pieces.Count = 0 Then GoTo NextWeek

        ' Sort latest-first across week: dateSortKey DESC, then DateTimeOut DESC
        Set pieces = SortedPiecesLatestFirst(pieces)

        ' Track how much each piece received in each band (for audit)
        Set f2Taken = CreateObject("Scripting.Dictionary")
        Set a2Taken = CreateObject("Scripting.Dictionary")
        Set baseTaken = CreateObject("Scripting.Dictionary")

        ' === Allocate per piece using max(daily, weekly) per band ===
        For i = 1 To pieces.Count
            With pieces(i)
                ptr = CStr(ObjPtr(pieces(i)))
                workDate = CDate(.data.DateIn)
                dow = .data.dayOfWeek

                ' How many minutes this piece "wants" in each band = union(daily, weekly) via max
                wantA2 = Application.WorksheetFunction.Max( _
                            DictGetLong(dailyA2, ptr), _
                            DictGetLong(weeklyA2, ptr))
                wantF2 = Application.WorksheetFunction.Max( _
                            DictGetLong(dailyF2, ptr), _
                            DictGetLong(weeklyF2, ptr))

                ' --- Band 2 (AFTER first 2h) first ---
                If .normalAvail > 0 And wantA2 > 0 Then
                    take = Application.WorksheetFunction.Min(.normalAvail, wantA2)
                    If take > 0 Then
                        ok = StoreWithCode(outOT, outNormal, .data, take, MapCodeForBand_DOW(dow, 2))
                        If ok Then
                            If Not a2Taken.Exists(ptr) Then a2Taken.Add ptr, 0&
                            a2Taken(ptr) = a2Taken(ptr) + take
                            .normalAvail = .normalAvail - take
                        'Else
                        '    LogError "Combine", .data.EmployeeCode, .data.DateIn, _
                        '             "A2StoreFailed", "Suppressed minutes in combiner", _
                        '             "Units=" & take & "; Code=" & MapCodeForBand_DOW(dow, 2)
                        End If
                    End If
                End If

                ' --- Band 1 (FIRST 2h) ---
                If .normalAvail > 0 And wantF2 > 0 Then
                    take = Application.WorksheetFunction.Min(.normalAvail, wantF2)
                    If take > 0 Then
                        ok = StoreWithCode(outOT, outNormal, .data, take, MapCodeForBand_DOW(dow, 1))
                        If ok Then
                            If Not f2Taken.Exists(ptr) Then f2Taken.Add ptr, 0&
                            f2Taken(ptr) = f2Taken(ptr) + take
                            .normalAvail = .normalAvail - take
                        'Else
                        '    LogError "Combine", .data.EmployeeCode, .data.DateIn, _
                        '             "F2StoreFailed", "Suppressed minutes in combiner", _
                        '             "Units=" & take & "; Code=" & MapCodeForBand_DOW(dow, 1)
                        End If
                    End If
                End If

                ' --- Base (whatever remains) ---
                If .normalAvail > 0 Then
                    take = .normalAvail
                    ok = StoreWithCode(outOT, outNormal, .data, take, MapCodeForBand_DOW(dow, 0))
                    If ok Then
                        If Not baseTaken.Exists(ptr) Then baseTaken.Add ptr, 0&
                        baseTaken(ptr) = baseTaken(ptr) + take
                        .normalAvail = 0
                    'Else
                    '    LogError "Combine", .data.EmployeeCode, .data.DateIn, _
                    '             "BaseStoreFailed", "Suppressed minutes in combiner", _
                    '             "Units=" & take & "; Code=" & MapCodeForBand_DOW(dow, 0)
                    End If
                End If
            End With
        Next i

        ' === Per-piece audit summary ===
        Dim rec(1 To 12) As Variant
        Dim sp As ShiftPieceClass

        For i = 1 To pieces.Count
            Set sp = pieces(i)
            workDate = CDate(sp.data.DateIn)
            dow = sp.data.dayOfWeek
            ptr = CStr(ObjPtr(sp))

            rec(1) = sp.data.EmployeeCode
            rec(2) = workDate
            rec(3) = sp.payrollExportCode
            rec(4) = sp.data.PayClassCode
            rec(5) = sp.data.companyCode    ' header says "CompanyCode/CostCentre" - you can join if you wish
            rec(6) = CLng(IIf(baseTaken.Exists(ptr), baseTaken(ptr), 0&))  ' BaseUnits
            rec(7) = CLng(IIf(f2Taken.Exists(ptr), f2Taken(ptr), 0&))      ' First2Units
            rec(8) = CLng(IIf(a2Taken.Exists(ptr), a2Taken(ptr), 0&))      ' After2Units
            rec(9) = MapCodeForBand_DOW(dow, 0)   ' BaseCode
            rec(10) = MapCodeForBand_DOW(dow, 1)  ' First2Code
            rec(11) = MapCodeForBand_DOW(dow, 2)  ' After2Code
            rec(12) = sp.data.weekSortKey

            bandWins(ptr) = rec
        Next i

NextWeek:
    Next wk
End Sub

' ---- Small helper: get 0 if key missing ----
Private Function DictGetLong(ByVal dict As Object, ByVal key As String) As Long
    If dict Is Nothing Then Exit Function
    If dict.Exists(key) Then
        DictGetLong = CLng(dict(key))
    Else
        DictGetLong = 0&
    End If
End Function
Private Sub SortPiecesLatestFirst(ByRef pcs As Collection)
    Dim i As Long, j As Long
    If pcs Is Nothing Or pcs.Count <= 1 Then Exit Sub
    For i = 2 To pcs.Count
        Dim cur As ShiftPieceClass
        Set cur = pcs(i)
        j = i - 1
        Do While j >= 1 And _
            (pcs(j).dateSortKey < cur.dateSortKey Or _
            (pcs(j).dateSortKey = cur.dateSortKey And pcs(j).DateTimeOut < cur.DateTimeOut))
            Dim tmp As ShiftPieceClass
            Set tmp = pcs(j)
            Set pcs(j) = cur
            Set pcs(j + 1) = tmp
            j = j - 1
        Loop
    Next i
End Sub



Private Sub WriteBandWinsAudit(ByVal bandWins As Object)
    Dim ws As Worksheet, r As Long, k As Variant, rec As Variant
    Set ws = GetOrCreateSheet("Audit_BandWins")
    ws.Cells.Clear

    ' Headers
    ws.Range("A1").value = "EmployeeCode"
    ws.Range("B1").value = "WorkDate"
    ws.Range("C1").value = "PayrollExportCode"
    ws.Range("D1").value = "PayClassCode"
    ws.Range("E1").value = "CompanyCode/CostCentre"
    ws.Range("F1").value = "BaseUnits"
    ws.Range("G1").value = "First2Units"
    ws.Range("H1").value = "After2Units"
    ws.Range("I1").value = "BaseCode"
    ws.Range("J1").value = "First2Code"
    ws.Range("K1").value = "After2Code"
    ws.Range("L1").value = "WeekSortKey"

    r = 2
    For Each k In bandWins.keys
        rec = bandWins(k)
        ws.Cells(r, 1).value = rec(1)
        ws.Cells(r, 2).value = rec(2)
        ws.Cells(r, 3).value = rec(3)
        ws.Cells(r, 4).value = rec(4)
        ws.Cells(r, 5).value = rec(5)
        ws.Cells(r, 6).value = rec(6)
        ws.Cells(r, 7).value = rec(7)
        ws.Cells(r, 8).value = rec(8)
        ws.Cells(r, 9).value = rec(9)
        ws.Cells(r, 10).value = rec(10)
        ws.Cells(r, 11).value = rec(11)
        ws.Cells(r, 12).value = rec(12)
        r = r + 1
    Next k

    ws.Columns.AutoFit
End Sub
' Allocate tail minutes within a single week's days (dates descending, then latest piece first)
' Returns minutesByPtr(ObjPtr(sp)) = minutes taken
Private Function AllocateFromLatestInWeekBucket(ByVal daysDict As Object, ByVal need As Long) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If need <= 0 Or daysDict Is Nothing Then Set AllocateFromLatestInWeekBucket = dict: Exit Function

    ' Build sortable array: arr(:,1)=yyyymmdd as Long, arr(:,2)=dayKey
    Dim arr() As Variant, c As Long
    Dim dayKey As Variant, col As Collection, dWork As Date
    ReDim arr(1 To IIf(daysDict Is Nothing, 0, daysDict.Count), 1 To 2)
    c = 0
    For Each dayKey In daysDict.keys
        Set col = daysDict(dayKey)
        If Not col Is Nothing And col.Count > 0 Then
            dWork = CDate(col(1).data.DateIn)
            c = c + 1
            arr(c, 1) = CLng(Format(dWork, "yyyymmdd"))
            arr(c, 2) = dayKey
        End If
    Next dayKey
    If c = 0 Then Set AllocateFromLatestInWeekBucket = dict: Exit Function
    
    If c < UBound(arr, 1) Then ReDim Preserve arr(1 To c, 1 To 2)

    ' Sort by date descending
    Dim i As Long, j As Long, t1 As Variant, t2 As Variant
    For i = 1 To c - 1
        For j = i + 1 To c
            If CLng(arr(j, 1)) > CLng(arr(i, 1)) Then
                t1 = arr(i, 1): arr(i, 1) = arr(j, 1): arr(j, 1) = t1
                t2 = arr(i, 2): arr(i, 2) = arr(j, 2): arr(j, 2) = t2
            End If
        Next j
    Next i

    ' Allocate from latest day, latest piece first (reverse iteration)
    Dim needLeft As Long: needLeft = need
    Dim sp As ShiftPieceClass, take As Long, key As String, m As Long

    For i = 1 To c
        If needLeft <= 0 Then Exit For
        dayKey = arr(i, 2)
        Set col = daysDict(dayKey)
        If col Is Nothing Then GoTo NextDay

        For m = col.Count To 1 Step -1
            If needLeft <= 0 Then Exit For
            Set sp = col(m)
            If sp.normalAvail > 0 Then
                take = CLng(Application.WorksheetFunction.Min(sp.normalAvail, needLeft))
                If take > 0 Then
                    key = CStr(ObjPtr(sp))
                    If Not dict.Exists(key) Then dict.Add key, 0&
                    dict(key) = dict(key) + take
                    needLeft = needLeft - take
                End If
            End If
        Next m
NextDay:
    Next i

    Set AllocateFromLatestInWeekBucket = dict
    DumpDictionary AllocateFromLatestInWeekBucket, "AllocateFromLatestInWeekBucket", 50
End Function


Private Sub LogOvertimeLeak(rowData As PayrollRowDataClass, ByVal normalUnits As Double)
    If IsOvertimeCode(rowData.payrollCode) Then Exit Sub
    If normalUnits > 11.5 Then
        If ErrorSheet Is Nothing Then Exit Sub
        Dim r As Long
        r = ErrorSheet.Cells(ErrorSheet.Rows.Count, "A").End(xlUp).Row + 1
        With ErrorSheet
            .Cells(r, 1).value = Now
            .Cells(r, 2).value = "NormalLeak"
            .Cells(r, 3).value = rowData.EmployeeCode
            .Cells(r, 4).value = rowData.DateIn
            .Cells(r, 5).value = "NormalTime > 11.5h"
            .Cells(r, 6).value = "PayRate=" & rowData.payRate
            .Cells(r, 7).value = "Units=" & Format(normalUnits, "0.0000")
        End With
    End If
End Sub
Function ParseYYMMDDToYYYYMMDD(ByVal yyMMDDString As String) As String
    Dim yearPart As String
    Dim monthPart As String
    Dim dayPart As String
    Dim fullDate As Date

    ' Extract year, month, and day parts from the YYMMDD string
    yearPart = Left(yyMMDDString, 2)
    monthPart = Mid(yyMMDDString, 3, 2)
    dayPart = Right(yyMMDDString, 2)

    ' Construct a full year (assuming 2000s for years 00-99)
    ' You might need to adjust the century logic based on your specific data
    If CInt(yearPart) >= 0 And CInt(yearPart) <= 99 Then ' Assuming 2000-2099
        yearPart = "20" & yearPart
    Else ' Handle other century if needed, e.g., "19" & yearPart for 90s
        ' Add more robust logic here if your data spans multiple centuries
    End If

    ' Create a Date object using DateSerial
    fullDate = DateSerial(CInt(yearPart), CInt(monthPart), CInt(dayPart))

    ' Format the Date object to YYYYMMDD string
    ParseYYMMDDToYYYYMMDD = Format(fullDate, "yyyymmdd")
End Function

Private Function IsBaseCode(code As String) As Boolean
    Select Case Trim$(code)
        Case "600", "601", "602": IsBaseCode = True
    End Select
End Function

Private Function IsOvertimeCode(code As String) As Boolean
    Select Case Trim$(code)
        Case "605", "606", "607", "608", "610", "603": IsOvertimeCode = True
    End Select
End Function

' Build a stable key to identify a shift across phases (Emp + WorkDate + Site)
Private Function BuildErrorShiftKey(ByRef d As PayrollRowDataClass) As String
    BuildErrorShiftKey = Trim$(d.EmployeeCode) & "|" & CStr(d.dateSortKey) & "|" & Trim$(d.payrollExportCode)
End Function

' Raw helper for allowances path (since ProcessAllowances derives from InputSheet)
Private Function BuildErrorShiftKeyRaw(empCode As String, dateSortKey As Long, payrollExportCode As String) As String
    BuildErrorShiftKeyRaw = Trim$(empCode) & "|" & CStr(dateSortKey) & "|" & Trim$(payrollExportCode)
End Function

' Units (hours*10000) -> Hours (double)
Private Function UnitsToHours(ByVal units As Double) As Double
    UnitsToHours = Round(units / 10000#, 4)
End Function

' Mark a shift as error and log with hours
Private Sub MarkErrorShift(ByRef d As PayrollRowDataClass, ByVal units As Double, ByVal reason As String)
    Dim key As String: key = BuildErrorShiftKey(d)
    If Not errorShiftKeys.Exists(key) Then errorShiftKeys.Add key, True
    ' Log with Hours and context = StoreWithCode
    Dim hrs As Double: hrs = UnitsToHours(units)
    Call LogErrorWithHours("StoreWithCode", d.EmployeeCode, d.DateIn, hrs, "StoreWithCode", "Validation failed", _
                           "Reason=" & reason & "; PayClass=" & d.PayClassCode & "; PayCode=" & d.payrollCode & _
                           "; Site=" & d.payrollExportCode & "; Co=" & d.companyCode)
End Sub

' Check if a given row belongs to an error-marked shift
Private Function IsErrorShift(ByRef d As PayrollRowDataClass) As Boolean
    Dim key As String: key = BuildErrorShiftKey(d)
    IsErrorShift = errorShiftKeys.Exists(key)
End Function


Public Sub DumpDictionary(ByVal dict As Object, _
                          Optional ByVal name As String = "dict", _
                          Optional ByVal maxItems As Long = 100)
    Dim k As Variant, n As Long

    If dict Is Nothing Then
        Debug.Print name & " <Nothing>"
        Exit Sub
    End If

    ' Accept late-bound Scripting.Dictionary as well
    Dim t As String: t = TypeName(dict)
    If InStr(1, t, "Dictionary", vbTextCompare) = 0 Then
        Debug.Print name & " is not a Dictionary (Type=" & t & ")"
        Exit Sub
    End If

    Debug.Print "=== Dump " & name & " (Count=" & dict.Count & ", Type=" & t & ") ==="
    For Each k In dict.keys
        n = n + 1
        If n > maxItems Then
            Debug.Print "� (" & (dict.Count - maxItems) & " more items not shown)"
            Exit For
        End If

        Debug.Print "Key: " & CStr(k) & " | ValueType: " & TypeName(dict(k))

        ' Handle common value types
        If TypeName(dict(k)) = "Collection" Then
            Dim col As Collection
            Dim i As Long
            Set col = dict(k)
            Debug.Print "  (Collection Count=" & col.Count & ")"
            For i = 1 To col.Count
                If TypeName(col(i)) = "ShiftPieceClass" Then
                    DumpShiftPiece col(i), "    [" & i & "] "
                Else
                    Debug.Print "    [" & i & "] " & TypeName(col(i))
                End If
            Next

        ElseIf InStr(1, TypeName(dict(k)), "Dictionary", vbTextCompare) > 0 Then
            ' Nested dictionary: recurse with a reduced max to avoid floods
            DumpDictionary dict(k), name & "->" & CStr(k), maxItems

        Else
            ' Simple scalar value (number/string/date)
            Debug.Print "  = "; dict(k)
        End If
    Next
End Sub

Public Sub DumpShiftPiece(ByVal sp As ShiftPieceClass, Optional ByVal prefix As String = "")
    On Error Resume Next ' Defensive: don�t interrupt dumps on a bad member
    Debug.Print prefix & _
        "Emp=" & sp.EmployeeCode & _
        "; Site=" & sp.payrollExportCode & _
        "; Date=" & Format$(sp.data.DateIn, "dd-mmm-yyyy") & _
        "; raw=" & Format$(sp.rawUnits / 10000#, "0.00") & _
        "; r1=" & Format$(sp.r1Units / 10000#, "0.00") & _
        "; avail=" & Format$(sp.normalAvail / 10000#, "0.00") & _
        "; PayClass=" & sp.data.PayClassCode & _
        "; Code=" & sp.data.payrollCode
End Sub
Public Sub DumpCollection(ByVal col As Collection, _
                          Optional ByVal label As String = "Collection")
    Dim i As Long, t As String
    Debug.Print "=== " & label & " (Count=" & col.Count & ") ==="
    For i = 1 To col.Count
        t = TypeName(col(i))
        Debug.Print "  [" & i & "] Type=" & t

        Select Case t
            Case "ShiftPieceClass":   DumpShiftPiece col(i), "    "
            Case "PayrollRowDataClass": DumpRowData col(i), "    "
            Case Else
                ' For scalars (String/Double/Date) you can print the value:
                If IsObject(col(i)) Then
                    Debug.Print "    (object: " & t & ")"
                Else
                    Debug.Print "    Value="; col(i)
                End If
        End Select
    Next
End Sub

Public Sub DumpRowData(ByVal d As PayrollRowDataClass, Optional ByVal prefix As String = "")
    On Error Resume Next
    Debug.Print prefix & _
        "Emp=" & d.EmployeeCode & _
        "; Co=" & d.companyCode & _
        "; Date=" & Format$(d.DateIn, "dd-mmm-yyyy") & _
        "; Code=" & d.payrollCode & _
        "; PayClass=" & d.PayClassCode & _
        "; Hours(units)=" & CLng(d.hours) & _
        "; CostCentre=" & d.costCentre
End Sub

' ==== NEW: Day-of-week based mapper to remove date/locale ambiguity ====
Private Function MapCodeForBand_DOW(ByVal dayOfWeek As Integer, ByVal band As Integer) As String
    ' band: 0=base, 1=first2, 2=after2
    Select Case dayOfWeek
        Case 1 To 5  ' Mon�Fri
            If band = 2 Then
                MapCodeForBand_DOW = "606"
            ElseIf band = 1 Then
                MapCodeForBand_DOW = "605"
            Else
                MapCodeForBand_DOW = "600"
            End If
        Case 6        ' Saturday
            If band = 2 Then
                MapCodeForBand_DOW = "608"
            ElseIf band = 1 Then
                MapCodeForBand_DOW = "607"
            Else
                MapCodeForBand_DOW = "601"
            End If
        Case Else     ' Sunday
            MapCodeForBand_DOW = "602"   ' base and both OT bands are 602 for Sunday in your scheme
    End Select
End Function

