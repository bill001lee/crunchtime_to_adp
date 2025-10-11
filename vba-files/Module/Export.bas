Attribute VB_Name = "Export"
'Sub Export()
'    Dim wsDataOut As Worksheet
'    Dim wsDOtimeOut As Worksheet
'    Dim wsAllowancesOut As Worksheet
'    Dim wsTemp As Worksheet
'    Dim exportPath As String
'    Dim companyCode As String
'    Dim exportDate As String
'    Dim fileName As String
'    Dim lastRowDataOut As Long
'    Dim lastRowOTDeduped As Long
'    Dim lastRowAllowancesOut As Long
'    Dim lastColDataOut As Long
'    Dim lastColOTDeduped As Long
'    Dim lastColAllowancesOut As Long
'    Dim lastRowTemp As Long
'    Dim i As Long, j As Long
'    Dim cellValue As String
'    Dim fileNum As Integer
'    Dim line As String
'
'    Set wsDataOut = ThisWorkbook.Sheets("NormalTime")
'    Set wsDOtimeOut = ThisWorkbook.Sheets("OTDeduped")
'    Set wsAllowancesOut = ThisWorkbook.Sheets("AllowancesOut")
'
'    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
'    wsTemp.name = "TempSheet"
'
'    wsTemp.Columns(2).NumberFormat = "@"
'    wsTemp.Columns(4).NumberFormat = "@"
'    wsTemp.Columns(5).NumberFormat = "@"
'    wsTemp.Columns(6).NumberFormat = "@"
'
'    companyCode = wsDataOut.Cells(2, 1).value
'    exportDate = Format(Date, "YYYYMMDD")
'    fileName = "paymast.dat"
'    exportPath = "C:\ADP\" & fileName
'
'    If Dir("C:\ADP\", vbDirectory) = "" Then MkDir "C:\ADP\"
'
'    lastRowDataOut = wsDataOut.Cells(wsDataOut.Rows.Count, "A").End(xlUp).Row
'    lastColDataOut = wsDataOut.Cells(1, wsDataOut.Columns.Count).End(xlToLeft).Column
'    lastRowOTDeduped = wsDOtimeOut.Cells(wsDOtimeOut.Rows.Count, "A").End(xlUp).Row
'    lastColOTDeduped = wsDOtimeOut.Cells(1, wsDOtimeOut.Columns.Count).End(xlToLeft).Column
'    lastRowAllowancesOut = wsAllowancesOut.Cells(wsAllowancesOut.Rows.Count, "A").End(xlUp).Row
'    lastColAllowancesOut = wsAllowancesOut.Cells(1, wsAllowancesOut.Columns.Count).End(xlToLeft).Column
'
'    wsDataOut.Range("A1").Resize(lastRowDataOut, lastColDataOut).Copy wsTemp.Range("A1")
'    wsDOtimeOut.Range("A2").Resize(lastRowOTDeduped - 1, lastColOTDeduped).Copy wsTemp.Range("A" & lastRowDataOut + 1)
'    wsAllowancesOut.Range("A2").Resize(lastRowAllowancesOut - 1, lastColAllowancesOut).Copy wsTemp.Range("A" & lastRowDataOut + lastRowOTDeduped)
'
'    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row
'
'    wsTemp.Sort.SortFields.Clear
'    wsTemp.Range("A1").Resize(lastRowTemp, lastColDataOut).Sort _
'        Key1:=wsTemp.Range("B1"), Order1:=xlAscending, _
'        Key2:=wsTemp.Range("L1"), Order2:=xlAscending, _
'        Key3:=wsTemp.Range("M1"), Order3:=xlAscending, _
'        Header:=xlYes
'
'    fileNum = FreeFile
'    Open exportPath For Output As #fileNum
'
'    For i = 2 To lastRowTemp
'        line = ""
'        For j = 1 To 11 'lastColDataOut - 2  ' Don't ecport the sort columns
'            cellValue = wsTemp.Cells(i, j).value
'            line = line & cellValue
'            If j < 11 Then line = line & ","
'        Next j
'        Print #fileNum, line
'    Next i
'
'
'    Close #fileNum
'
'    Application.DisplayAlerts = False
'    wsTemp.Delete
'    Application.DisplayAlerts = True
'
'    MsgBox "Data exported successfully to C:\ADP\", vbInformation
'End Sub

Public Sub Export()
    Dim wsTemp As Worksheet
    Dim tmpName As String: tmpName = "TempSheet"
    Dim exportPath As String
    Dim fDialog As fileDialog
    Dim lastRow As Long, pasteRow As Long

    On Error GoTo EH

    ' Delete existing TempSheet if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(tmpName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Add new TempSheet
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsTemp.name = tmpName

    ' Headers
    wsTemp.Range("A1:K1").value = Array("OwnershipEntity", "PayrollExportCode", "WeekEndingDate", _
                                        "PayrollID", "EmployeePositionCode", "GLNumber", _
                                        "DateIn", "DateOut", "TimeIn", "TimeOut", "PayRate")

    pasteRow = 2

    ' Copy NormalTime
    lastRow = Sheets("NormalTime").Cells(Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        Sheets("NormalTime").Range("A2:K" & lastRow).Copy wsTemp.Range("A" & pasteRow)
        pasteRow = pasteRow + (lastRow - 1)
    End If

    ' Copy OTDeduped
    lastRow = Sheets("OTDeduped").Cells(Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        Sheets("OTDeduped").Range("A2:K" & lastRow).Copy wsTemp.Range("A" & pasteRow)
        pasteRow = pasteRow + (lastRow - 1)
    End If

    ' Copy AllowancesOut
    lastRow = Sheets("AllowancesOut").Cells(Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        Sheets("AllowancesOut").Range("A2:K" & lastRow).Copy wsTemp.Range("A" & pasteRow)
        pasteRow = pasteRow + (lastRow - 1)
    End If

    ' Sort by Employee, Week, Date
    wsTemp.Sort.SortFields.Clear
    wsTemp.Sort.SortFields.Add key:=wsTemp.Range("B2:B" & pasteRow), Order:=xlAscending
    wsTemp.Sort.SortFields.Add key:=wsTemp.Range("L2:L" & pasteRow), Order:=xlAscending
    wsTemp.Sort.SortFields.Add key:=wsTemp.Range("M2:M" & pasteRow), Order:=xlAscending
    With wsTemp.Sort
        .SetRange wsTemp.Range("A1:M" & pasteRow)
        .Header = xlYes
        .Apply
    End With

    ' Ask user for export folder
    Set fDialog = Application.fileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select Export Folder"
        If .show <> -1 Then
            MsgBox "Export cancelled.", vbExclamation
            GoTo CleanExit
        End If
        exportPath = .SelectedItems(1) & "\paymast.dat"
    End With

    ' Write to CSV file
    Dim fileNum As Integer
    fileNum = FreeFile
    Open exportPath For Output As #fileNum
    Dim r As Range
    For Each r In wsTemp.Range("A1:K" & pasteRow).Rows
        Print #fileNum, Join(Application.Transpose(Application.Transpose(r.value)), ",")
    Next r
    Close #fileNum

    MsgBox "Export complete: " & exportPath, vbInformation

CleanExit:
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True
    Exit Sub

EH:
    MsgBox "Error during export: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

