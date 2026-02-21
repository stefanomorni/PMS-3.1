VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHFM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
SAAProc.PorfolioSheetSelectionChange Target
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
SAAProc.PortfolioSheetChange Target
End Sub
Private Sub Aggiorna_Campi_MarketScreener(ByVal Caller As Object)
    ISIN = Caller.Text
    MainURL = "https://www.marketscreener.com/"
    Text_To_Search = scrapehtml.GetPageContent("https://www.marketscreener.com/search/instruments?q=" & ISIN)

    Field_Code = "MS_Code"
    SearchPattern = "(?:codezb=\x22)(.+?)\x22 at"
    Search_Result = morfunctions.RegExMatch(Text_To_Search, SearchPattern)
    If Search_Result <> "" Then
        MS_Code = Search_Result
    End If
    
    Field_Code = "MS_W"
    SearchPattern = "(?:ALNI0[\s\S]+?href=\x22/)(.+?)\x22 codezb"
    Set Target_Cell = Caller.Parent.Cells(Caller.Row, Caller.Parent.Range("rwIntestazioni").Cells.Find(Field_Code).Column)
    Search_Result = morfunctions.RegExMatch(Text_To_Search, SearchPattern)
    Prefix = MainURL
    If Search_Result <> "" Then
        Hyperlink_Formula = Prefix & Search_Result
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
    Prefix = MainURL
    If MS_Code <> "" Then
        Hyperlink_Formula = Prefix & "mods_a/charts/TV/inc_dyna_graph.php?c=" & MS_Code & "&h=950&w=1900&style=TradingView.CANDLES&interval=(null)&mobile=(null)&showAT=1"
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
Private Sub AggiornadatiFoglio()
SAAProc.AggiornadatiSAA
End Sub
Private Sub OrdineTrasmesso()
SAAProc.AggiornaPosizione Selection
End Sub
Private Sub RimuoviTitolo()
SAAProc.RimuoviTitolo Selection
End Sub
Private Sub AggiungiTitoloSopra()
SAAProc.AggiungiTitolo Selection, "Sopra"
End Sub
Private Sub AggiungiTitoloSotto()
SAAProc.AggiungiTitolo Selection, "Sotto"
End Sub
Private Sub EliminaOrdini()
SAAProc.EliminaOrdini Selection
End Sub
Private Sub Esporta_Ordini()
SAAProc.Esporta_Ordini Selection
End Sub
Private Sub Registra_Ordini()
SAAProc.Registra_Ordini Selection
End Sub
Private Sub ResetLukbId()
SAAProc.ResetLukbId Selection
End Sub
Private Sub Worksheet_Activate()
FoglioAttivo = ActiveSheet.Name
AggiungiComandoAMenu Me, "Cell", "Aggiorna Prezzi", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".AggiornaDatiFoglio")
AggiungiComandoAMenu Me, "Cell", "Rimuovi Titolo", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".RimuoviTitolo")
AggiungiComandoAMenu Me, "Cell", "Aggiungi Titolo Sotto", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".AggiungiTitoloSotto")
AggiungiComandoAMenu Me, "Cell", "Aggiungi Titolo Sopra", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".AggiungiTitoloSopra")
AggiungiComandoAMenu Me, "Cell", "Elimina Ordini", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".EliminaOrdini")
AggiungiComandoAMenu Me, "Cell", "Esporta Ordini", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".Esporta_Ordini")
AggiungiComandoAMenu Me, "Cell", "Registra Ordini", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".Registra_Ordini")
AggiungiComandoAMenu Me, "Cell", "Ordine(i) Trasmesso(i)", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".OrdineTrasmesso")
AggiungiComandoAMenu Me, "Cell", "Reset fonte prezzo", 1, VBA.Trim(" '" & Me.Parent.Name & "'!" & Me.CodeName & ".ResetLukbId")

End Sub

Private Sub Worksheet_Deactivate()
EliminaComandoDaMenu "Cell", "Aggiorna Prezzi"
EliminaComandoDaMenu "Cell", "Ordine(i) Trasmesso(i)"
EliminaComandoDaMenu "Cell", "Rimuovi Titolo"
EliminaComandoDaMenu "Cell", "Elimina Ordini"
EliminaComandoDaMenu "Cell", "Esporta Ordini"
EliminaComandoDaMenu "Cell", "Registra Ordini"
EliminaComandoDaMenu "Cell", "Reset fonte prezzo"
End Sub


