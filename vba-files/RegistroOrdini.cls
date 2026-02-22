VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistroOrdini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Application.ScreenUpdating = False
Application.EnableEvents = False
'Sort
With RegistroOrdini.ListObjects("Registro_Ordini")
    .Sort.SortFields.Clear
    .Sort.SortFields.Add2 _
        Key:=Range("Registro_Ordini[Data]"), SortOn:=xlSortOnValues, Order:=xlDescending
    .Sort.SortFields.Add2 _
        Key:=Range("Registro_Ordini[Ora]"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortTextAsNumbers
    .Sort.SortFields.Add2 _
        Key:=Range("Registro_Ordini[Titolo]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers
    .Sort.SortFields.Add2 _
        Key:=Range("Registro_Ordini[Cliente]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers
    With .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End With

Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub
