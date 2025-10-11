Attribute VB_Name = "Import"
Sub Import()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileDialog As fileDialog
    Dim lastRow As Long
    On Error GoTo ErrorHandler

    ' Target sheet
    Set ws = ThisWorkbook.Sheets("DataIn")
    ws.Cells.Clear

    ' Headers
    ws.Range("A1:K1").value = Array("OwnershipEntity", "PayrollExportCode", "WeekEndingDate", _
                                    "PayrollID", "EmployeePositionCode", "GLNumber", _
                                    "DateIn", "DateOut", "TimeIn", "TimeOut", "PayRate")

    ' File picker
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select CSV File"
        .InitialFileName = "C:\ADP\"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        If .show = -1 Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "No file selected.", vbExclamation
            Exit Sub
        End If
    End With

    ' Import as CSV (comma-delimited)
    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A2"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileConsecutiveDelimiter = False
        .TextFileTrailingMinusNumbers = True
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .Refresh BackgroundQuery:=False
    End With

    ' Optional: Auto-fit
    ws.Columns.AutoFit
    MsgBox "Data imported successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

