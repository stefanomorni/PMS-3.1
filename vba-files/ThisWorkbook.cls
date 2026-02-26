VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("PMS.Classes")
Dim Calcolazione
Private Sub Workbook_Activate()
Calcolazione = Application.Calculation
Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub Workbook_Deactivate()
Application.Calculation = Calcolazione
On Error Resume Next
ActiveWorkbook.Connections("Query - Posizioni").OLEDBConnection.RefreshPeriod = 0
End Sub

Private Sub Workbook_Open()
End Sub
