VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Variable to store the previous calculation state
Private mPrevCalculation As XlCalculation

Private Sub Workbook_Activate()
    On Error Resume Next
    mPrevCalculation = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    SetRefreshPeriod "Query - Posizioni", 1
    On Error GoTo 0
End Sub

Private Sub Workbook_Deactivate()
    ' Restore calculation if it was set
    On Error Resume Next
    If mPrevCalculation <> 0 Then
        Application.Calculation = mPrevCalculation
    End If
    SetRefreshPeriod "Query - Posizioni", 0
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    SetRefreshPeriod "Query - Posizioni", 1
End Sub

''' <summary>
''' Safely sets the RefreshPeriod of a connection if it exists and supports OLEDB.
''' Prevents Error 91 if the connection is missing or not initialized.
''' </summary>
Private Sub SetRefreshPeriod(ByVal connName As String, ByVal period As Long)
    On Error Resume Next
    Dim conn As WorkbookConnection
    Set conn = ThisWorkbook.Connections(connName)
    
    If Not conn Is Nothing Then
        conn.OLEDBConnection.RefreshPeriod = period
    End If
    On Error GoTo 0
End Sub
