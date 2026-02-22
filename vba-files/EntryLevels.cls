Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim Table As Object
If Selection.Row > 2 Then
    Set Table = morfunctions.ActiveTable(Target)
    MORProcedures.Highlight_Selected_Tabe_Row Table, Target
End If
End Sub
