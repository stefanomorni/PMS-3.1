VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
'Test comment
Application.ScreenUpdating = False
With Me.ListObjects("Transazioni").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim Table As Object
If Selection.Row > 2 Then
    Set Table = morfunctions.ActiveTable(Target)
    MORProcedures.Highlight_Selected_Tabe_Row Table, Target
End If
End Sub
