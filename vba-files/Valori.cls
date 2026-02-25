VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Valori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public UserChange As Boolean
Public Sub Restore_Ambient()
UserChange = True
Application.ScreenUpdating = True
'Application.AutoCorrect.AutoFillFormulasInLists = True
End Sub

Private Sub Worksheet_Activate()
Application.ScreenUpdating = False
Application.EnableEvents = False

On Error GoTo fine:

FoglioAttivo = ActiveSheet.Name
AggiungiComandoAMenu Me, "List Range Popup", "Aggiorna Prezzi", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".AggiornaDatiFoglio")
AggiungiComandoAMenu Me, "List Range Popup", "Aggiorna Hyperlinks", 2, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".AggiornaHyperlinks")
AggiungiComandoAMenu Me, "List Range Popup", "Reset fonte prezzo", 2, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".ResetLukbId")


'Sort and Format
With Valori.ListObjects("Valori")
    .Sort.SortFields.Clear
    .Sort.SortFields.Add2 _
        Key:=Range("Valori[AC]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        CustomOrder:="CUR,BON,ECU,COM,ETF,FON", DataOption:=xlSortNormal
    .Sort.SortFields.Add2 _
        Key:=Range("Valori[Maturity]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    .Sort.SortFields.Add2 _
        Key:=Range("Valori[Descrizione]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With .Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End With

Range("Valori[Price_Link]").Font.Name = "Webdings"
Range("Valori[Price_Link]").Font.Underline = xlUnderlineStyleNone
Range("Valori[Info]").Font.Name = "Webdings"
Range("Valori[Info]").Font.Underline = xlUnderlineStyleNone
Range("Valori[FN_Link]").Font.Name = "smo-icons"
Range("Valori[FN_Link]").Font.Underline = xlUnderlineStyleNone
Range("Valori[MS_W]").Font.Name = "smo-icons"
Range("Valori[MS_W]").Font.Underline = xlUnderlineStyleNone
Range("Valori[Chart]").Font.Name = "smo-icons"
Range("Valori[Chart]").Font.Underline = xlUnderlineStyleNone
Selection.Offset(-1, 0).Select
Selection.Offset(1, 0).Select

fine:
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub
Private Sub Worksheet_Deactivate()
EliminaComandoDaMenu "List Range Popup", "Aggiorna Prezzi"
EliminaComandoDaMenu "List Range Popup", "Aggiorna Hyperlinks"
EliminaComandoDaMenu "List Range Popup", "Reset fonte prezzo"
End Sub

'Private Sub Worksheet_Change(ByVal Target As Range)
'On Error GoTo Reset
'If Target.Cells.Count > 1 Then Exit Sub
'If UserChange = True Then
'    'Temporay disable to avoid autofill by Hyperlink insertion
'    Application.AutoCorrect.AutoFillFormulasInLists = False
'    For Each cella In Target.Cells
'        If cella.Parent.Cells(cella.Parent.Range("rwIntestazioni").Row, cella.Column).Text = "smo_id" Then
'            If cella.Text <> "" Then
'                Aggiorna_Campi_MarketScreener cella
'                'Aggiorna_Campi_CapitalCube cella
'                ' Aggiorna_Campi_Finanzen_ch cella
'            End If
'        End If
'    Next cella
'End If
'Reset:
'    Restore_Ambient
'End Sub

Private Sub Aggiorna_Campi_TradingView(ByVal Caller As Object)

End Sub

Private Sub ResetLukbId()
SAAProc.ResetLukbId Selection
End Sub

Private Sub Aggiorna_Campi_MarketScreener(ByVal Caller As Object)
    UserChange = False
    ISIN = Caller.Text
    MainURL = "https://www.marketscreener.com/"
    Text_To_Search = scrapehtml.GetPageContent("https://www.marketscreener.com/search/instruments/?aComposeInputSearch=s_" & ISIN)

    Field_Code = "MS_Code"
    SearchPattern = "(?:codezb=\x22)(.+?)\x22 at"
    Set Target_Cell = Caller.Parent.Cells(Caller.Row, Caller.Parent.Range("rwIntestazioni").Cells.Find(Field_Code).Column)
    Search_Result = morfunctions.RegExMatch(Text_To_Search, SearchPattern)
    If Search_Result <> "" Then
        MS_Code = Search_Result
        Target_Cell.Formula = MS_Code
    End If

    Field_Code = "MS_Link"
    SearchPattern = "(?:ALNI0[\s\S]+?href=\x22/)(.+?)\x22 codezb"
    Set Target_Cell = Caller.Parent.Cells(Caller.Row, Caller.Parent.Range("rwIntestazioni").Cells.Find(Field_Code).Column)
    Search_Result = morfunctions.RegExMatch(Text_To_Search, SearchPattern)
    PREFIX = MainURL
    If Search_Result <> "" Then
        Hyperlink_Formula = PREFIX & Search_Result
        simbolo = Symbol("E903")
        Target_Cell.Formula2R1C1 = "=HYPERLINK(""" & Hyperlink_Formula & """,""" & simbolo & """)"
        With Target_Cell.Font
        .Name = "smo-icons"
        .Size = 9
        .Underline = xlNone
        .ThemeColor = xlThemeColorHyperlink
        End With
    Else
        Target_Cell.Formula = ""
    End If

    Field_Code = "MS_Chart"
    Set Target_Cell = Caller.Parent.Cells(Caller.Row, Caller.Parent.Range("rwIntestazioni").Cells.Find(Field_Code).Column)
    PREFIX = MainURL
    If MS_Code <> "" Then
        Hyperlink_Formula = PREFIX & "mods_a/charts/TV/inc_dyna_graph.php?c=" & MS_Code & "&h=950&w=1900&style=TradingView.CANDLES&interval=(null)&mobile=(null)&showAT=1"
        simbolo = Symbol("E934")
        Target_Cell.Formula2R1C1 = "=HYPERLINK(""" & Hyperlink_Formula & """,""" & simbolo & """)"
        With Target_Cell.Font
        .Name = "smo-icons"
        .Size = 9
        .Underline = xlNone
        .ThemeColor = xlThemeColorHyperlink
        End With
    Else
        Target_Cell.Formula = ""
    End If
    UserChange = True
End Sub
Private Sub Aggiorna_Campi_Finanzen_ch(ByVal Caller As Object)
    UserChange = False
    ISIN = Caller.Text
    Field_Code = "VN"
    Set Target_Cell = Caller.Parent.Cells(Caller.Row, Caller.Parent.Range("rwIntestazioni").Cells.Find(Field_Code).Column)
    VN = fin_ch_sec_info(ISIN, "valoren")
    If VN <> "" Then
        Target_Cell.value = VN
    End If

    Field_Code = "FN_Link"
    Set Target_Cell = Caller.Parent.Cells(Caller.Row, Caller.Parent.Range("rwIntestazioni").Cells.Find(Field_Code).Column)
    Url = fin_ch_sec_info(ISIN, "link")
    If Url <> "" Then
        Hyperlink_Formula = Url
        simbolo = Symbol("E903")
        Target_Cell.Formula2R1C1 = "=HYPERLINK(""" & Hyperlink_Formula & """,""" & simbolo & """)"
        With Target_Cell.Font
        .Name = "smo-icons"
        .Size = 9
        .Underline = xlNone
        .ThemeColor = xlThemeColorHyperlink
        End With
    Else
        Target_Cell.Formula = ""
    End If

    Field_Code = "MS_Chart"
    Set Target_Cell = Caller.Parent.Cells(Caller.Row, Caller.Parent.Range("rwIntestazioni").Cells.Find(Field_Code).Column)
    PREFIX = MainURL
    If MS_Code <> "" Then
        Hyperlink_Formula = PREFIX & "mods_a/charts/TV/inc_dyna_graph.php?c=" & MS_Code & "&h=950&w=1900&style=TradingView.CANDLES&interval=(null)&mobile=(null)&showAT=1"
        simbolo = Symbol("E934")
        Target_Cell.Formula2R1C1 = "=HYPERLINK(""" & Hyperlink_Formula & """,""" & simbolo & """)"
        With Target_Cell.Font
        .Name = "smo-icons"
        .Size = 9
        .Underline = xlNone
        .ThemeColor = xlThemeColorHyperlink
        End With
    Else
        Target_Cell.Formula = ""
    End If
    UserChange = True
End Sub
Private Function RegExMatch(Text_To_Search, RegExPattern, Optional Index As Integer = 0)
    Dim html As Object, objResult As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    RegEx.Pattern = RegExPattern
    RegEx.Global = True
    If Index < 0 Then Index = 0
    If RegEx.test(Text_To_Search) Then
        Set matches = RegEx.Execute(Text_To_Search)
        RegExMatch = matches(Index).SubMatches(0)
    Else
        RegExMatch = ""
    End If
End Function


Private Sub AggiornadatiFoglio()
SAAProc.AggiornadatiSAA
End Sub

Private Sub AggiornaHyperlinks()
MORProcedures.Aggiorna_Hyperlinks ActiveCell
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim Table As Object
If Selection.Row > 2 Then
    Set Table = morfunctions.ActiveTable(Target)
    MORProcedures.Highlight_Selected_Tabe_Row Table, Target
End If
End Sub

'-------------------------------------------------------------------------------
' Function: InWorksheets
'
' Description:
'   This function checks if the text of a given cell is present in any of the
'   worksheets in the current workbook.
'
' Parameters:
'   - sheetList (String): A string containing sheet names separated by a colon and space (e.g., "CHF M:EUR M").
'   - value (Range): A reference to the cell containing the text to search for.
'
' Returns:
'   - Integer: Returns 1 if the text is found in any worksheet, 0 otherwise.
'
' Example Usage:
'   Dim result As Integer
'   result = InWorksheets("CHF M:EUR M", Range("A1"))
'   MsgBox "Text found in worksheets: " & result
'
' Notes:
'   - Assumes that the cell value is a string.
'   - Case-insensitive search.
'-------------------------------------------------------------------------------
Function InWorksheets(sheetList As String, value As Range) As Integer
    Dim ws As Worksheet
    Dim cellValue As String
    Dim found As Boolean
    Dim sheetNames() As String
    Dim i As Long
    
    ' Initialize found flag
    found = False
    
    ' Split the sheetList string into an array of sheet names
    sheetNames = Split(sheetList, ":")
    
    ' Get the value of the cell
    cellValue = value.value
    
    ' Loop through each sheet name
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = Worksheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ' Check if the cell value is present in the worksheet
            If Not ws.Cells.Find(What:=cellValue, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False) Is Nothing Then
                found = True
                Exit For
            End If
        End If
    Next i
    
    ' Return 1 if found, 0 if not found
    If found Then
        InWorksheets = 1
    Else
        InWorksheets = 0
    End If
End Function

