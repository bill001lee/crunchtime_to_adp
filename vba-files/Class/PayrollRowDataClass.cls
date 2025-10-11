VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PayrollRowDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public companyCode As String
Public EmployeeCode As String
Public entryDate As String
Public payrollCode As String
Public PayClassCode As String
Public costCentre As String
Public fromDate As String
Public toDate As String
Public weekSortKey As Long
Public dateSortKey As Long
Public payrollExportCode As String
Public payRate As Variant
Public dayOfWeek As Integer
Public DateIn As Date
Public DateOut As Date
Public DateTimeIn As Date
Public DateTimeOut As Date
Public hours As Double
Public weekEnding As String
