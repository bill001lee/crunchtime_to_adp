Attribute VB_Name = "VersionInfo"
Option Explicit

' ==========================
' Translator build metadata
' ==========================
' Bump TRANSLATOR_VERSION when you cut a new coordinated change set
Public Const TRANSLATOR_VERSION As String = "2025.10.03"
' Comma-separated list of patch IDs you have applied in this build
' (edit this any time you change the applied patch set)
Public Const TRANSLATOR_PATCHES As String = _
    "T-2025-10-03-001,T-2025-10-03-002,T-2025-10-03-003,T-2025-10-03-004"

' ISO 8601 local build timestamp (optional but useful in logs)
Public Const TRANSLATOR_BUILD_STAMP As String = "2025-10-03T06:06:39+08:00"

' Human-readable one-liner for logs and splash messages
Public Property Get TranslatorVersionInfo() As String
    TranslatorVersionInfo = "Translator " & TRANSLATOR_VERSION & _
        " | patches=" & TRANSLATOR_PATCHES & _
        " | build=" & TRANSLATOR_BUILD_STAMP
End Property

' ==========================
' Patch list helpers
' ==========================
Public Function PatchList() As Variant
    PatchList = Split(TRANSLATOR_PATCHES, ",")
End Function

Public Function HasPatch(ByVal patchId As String) As Boolean
    Dim p As Variant
    For Each p In PatchList()
        If StrComp(Trim$(CStr(p)), Trim$(patchId), vbTextCompare) = 0 Then
            HasPatch = True
            Exit Function
        End If
    Next p
End Function

' ==========================
' Logging helper
' ==========================
' Call this at the end of TranslatePayroll (or ProcessPayroll) to stamp the build
Public Sub LogVersionToErrors(ByVal p As PayrollProcessor)
    On Error Resume Next
    If p Is Nothing Then Exit Sub
    If p.ErrorSheet Is Nothing Then Exit Sub

    With p.ErrorSheet
        Dim r As Long: r = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
        .Cells(r, 1).value = Now
        .Cells(r, 2).value = "Build"
        .Cells(r, 5).value = "Translator version"
        .Cells(r, 6).value = TranslatorVersionInfo
        .Rows(1).Font.Bold = True
    End With
    On Error GoTo 0
End Sub
'' Updated by Copilot on 2025-10-03 06:06:39 UTC+08:00
'' File: VersionInfo.bas
'' Source of truth for build/patch metadata.
'' End of file.
''
''
