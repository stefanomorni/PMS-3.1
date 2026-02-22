Private Sub Worksheet_Activate()
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
