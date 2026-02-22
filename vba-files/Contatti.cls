VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Contatti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Application.ScreenUpdating = False
With Contatti.ListObjects("Contatti")
    .Sort.SortFields.Clear
    .Sort.SortFields. _
        Add2 Key:=Range("B1:B54"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
End With
Range("A2").Select
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim Table As Object
If Selection.Row > 2 Then
    Set Table = morfunctions.ActiveTable(Target)
    MORProcedures.Highlight_Selected_Tabe_Row Table, Target
End If
End Sub
