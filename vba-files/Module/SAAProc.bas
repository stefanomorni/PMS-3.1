Attribute VB_Name = "SAAProc"
Option Explicit
'Declarations
Public FoglioAttivo As String
Public CurrentBook As Workbook
Public Procedure As String
Public SecondsFrequency As Single
Public Valore, Posizione As Range
Public ValoreSelezione
Public failed_updates As String
Public row_id As String
Public sec_id As String
Public calculation_setting
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
Public Function InWorksheets(sheetList As String, value As Range) As Integer
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


Sub Reset_Ambiente()
Attribute Reset_Ambiente.VB_ProcData.VB_Invoke_Func = "r\n14"
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

Sub Disable_Events()
Application.EnableEvents = False
Application.ScreenUpdating = True
End Sub

Sub Enable_Events()
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub

Sub EseguiOgniXSecondi(Optional Libro As Workbook, Optional Procedura As String, Optional SecondiFrequenza As Single)
'
' Runs the"procedure" every "Frequency" minutes from first call on the "CurrentBook" as long as it is still active; the "active check"
' every "frequency" minutes remains in action until the "Currentbook" is closed

Dim CurrentProcedure As String
Dim found As Boolean
Dim UpdateTime As Date
Dim BlockTime As Date

'Initialisations

Let CurrentProcedure = "EseguiOgniXSecondi"

'Procedure

If Procedura <> "" Then
    Let Procedure = Procedura
    Set CurrentBook = Libro
End If
On Error GoTo Filechiuso
'CurrentBook.Activate
Application.Run ("'" & CurrentBook.Name & "'!" & Procedure)
    'Tests values
    If IsMissing(SecondiFrequenza) Or SecondiFrequenza = 0 Then
        If SecondsFrequency < 60 Then SecondsFrequency = 60
    Else
        SecondsFrequency = SecondiFrequenza
    End If
     UpdateTime = Now() + TimeSerial(0, 0, SecondsFrequency)
     Application.OnTime EarliestTime:=UpdateTime, Procedure:=CurrentProcedure
Filechiuso:
End Sub
Sub Nascondi_Mostra_Proposta()
Attribute Nascondi_Mostra_Proposta.VB_ProcData.VB_Invoke_Func = "P\n14"
Switch_Visibilita Selection, "P_"
End Sub
Sub Nascondi_Mostra_Ordine()
Attribute Nascondi_Mostra_Ordine.VB_ProcData.VB_Invoke_Func = "O\n14"
Switch_Visibilita Selection, "O_"
End Sub
Sub Nascondi_Mostra_Attuale()
Attribute Nascondi_Mostra_Attuale.VB_ProcData.VB_Invoke_Func = "A\n14"
Switch_Visibilita Selection, "I_"
End Sub
Sub Nascondi_Mostra_Nuovo()
Attribute Nascondi_Mostra_Nuovo.VB_ProcData.VB_Invoke_Func = "N\n14"
Switch_Visibilita Selection, "S_"
End Sub

Sub Switch_Visibilita(Target, Prefix)
Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
On Error GoTo fine
Dim Cella As Range
For Each Cella In Intersect(Target.Parent.Range("rwIntestazioni"), Target.Parent.UsedRange)
    If Left(Cella, 2) = Prefix Then
            Cella.EntireColumn.Hidden = Not (Cella.EntireColumn.Hidden)
    End If
Next Cella
fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
End Sub
Sub Mostra_Tutto()
Attribute Mostra_Tutto.VB_ProcData.VB_Invoke_Func = "T\n14"
Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.ScreenUpdating = False
On Error GoTo fine
Dim Cella As Range
For Each Cella In Intersect(ActiveSheet.Range("rwIntestazioni"), ActiveSheet.UsedRange)
    If Mid(Cella, 2, 1) = "_" Then
            Cella.EntireColumn.Hidden = False
    End If
Next Cella
fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
End Sub

Sub ResetPosizioni(ByRef Target As Range)
Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
On Error GoTo fine
Dim Cella As Range
For Each Cella In Intersect(Target.EntireRow.Cells, Target.Parent.UsedRange)
    If Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "I_Pos" Or _
            Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "O_Pos" Or _
            Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "I_Price" Or _
            Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "Pct." Then
        Cella.ClearContents
        Cella.ClearComments
    End If
Next Cella
fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
'Application.Calculation = SAAProc.calculation_setting
End Sub
Sub RimuoviTitolo(ByRef Target As Range)
    Dim DataRow, ColonnaPosizione As Range
    Dim Posizione As Double
    Dim UserResponse As VbMsgBoxResult
    
    MORProcedures.CancelNamedQueryAndWait "Posizioni"
    
    Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
    Application.ScreenUpdating = False
    
    On Error GoTo fine
    
    Set DataRow = Target.EntireRow
    Dim rng_ColonnaPosizione As Range
Set rng_ColonnaPosizione = Target.Parent.Cells.Find(What:="Posizione", After:=Cells(1, 1), LookIn:=xlFormulas2, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, SearchFormat:=False)
If Not rng_ColonnaPosizione Is Nothing Then
Set ColonnaPosizione = rng_ColonnaPosizione.EntireColumn
End If

    Posizione = Intersect(DataRow, ColonnaPosizione).Value2
    
    If Posizione > 0 Then
        UserResponse = MsgBox("Sembra che ci siano dei portafogli con questo titolo." & Chr(10) & " Sei sicuro di volerlo eliminare?", vbOKCancel + vbExclamation, "Attenzione")
        
        If UserResponse = vbOK Then
            DataRow.Delete Shift:=xlUp
        Else
            GoTo fine
        End If
    Else
        DataRow.Delete Shift:=xlUp
    End If

fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
End Sub

Sub AggiungiTitolo(ByVal Target As Range, Posizione As String)
    Dim DataRow, ColonnaDescrizione As Range
    Dim xlShift
    
    Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
    Application.ScreenUpdating = False
    '' Temprary disabling calculation doesn't help.
    '' When reenabling, the all worksheet gets recalulated anyhow.
    'SAAProc.calculation_setting = Application.Calculation
    'Application.Calculation = xlCalculationManual
    
    On Error GoTo fine
    
    Dim rng_ColonnaDescrizione As Range
Set rng_ColonnaDescrizione = Target.Parent.Cells.Find(What:="Descrizione", After:=Cells(1, 1), LookIn:=xlFormulas2, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, SearchFormat:=False)
If Not rng_ColonnaDescrizione Is Nothing Then
Set ColonnaDescrizione = rng_ColonnaDescrizione.EntireColumn
End If

    Set DataRow = Target.EntireRow
    DataRow.Copy
    If Posizione = "Sopra" Then
        DataRow.Select
    Else
        DataRow.Offset(1, 0).Select
    End If
    Selection.Insert Shift:=xlDown
    Selection.SpecialCells(xlCellTypeConstants, 23).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Intersect(Selection, ColonnaDescrizione).Activate
fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
    'Application.Calculation = SAAProc.calculation_setting
End Sub
Sub EliminaOrdini(ByRef Target As Range)
Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.ScreenUpdating = False
On Error GoTo fine
Dim Cella As Range
For Each Cella In Intersect(Target.EntireRow.Cells, Target.Parent.UsedRange)
    If Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "O_Pos" Then
        Cella.ClearContents
        Cella.ClearComments
    End If
Next Cella
fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
End Sub
Sub Registra_Ordini(ByRef Target As Range)
Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.ScreenUpdating = False
On Error GoTo fine
Dim Cella As Range
For Each Cella In Intersect(Target.EntireRow.Cells, Target.Parent.UsedRange)
    If Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "O_Pos" And Cella.value <> "" Then
        'cella.Select
        AggiornaPosizione Cella.Cells(1, 1) 'Selection
    End If
Next Cella
fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
End Sub
Sub Esporta_Ordini(ByRef Target As Range)
Dim Ordini As Collection
Dim O As clsOrdine
Dim Calcolazione As XlCalculation
Dim wksName As String
Dim OrdersWs As Worksheet
Dim OrdersTbl As ListObject
Dim OrderRw As ListRow
Dim Security As String, ISIN As String, Crncy As String, Client As String, Portfolio As String
Dim Price As Double, I_Pos As Double, O_Pos As Double, Amnt As Double
Dim Valoren As String
Dim Ordine As clsOrdine
Dim Riga As Range

Calcolazione = Application.Calculation
Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
On Error GoTo fine
Set Ordini = New Collection

Dim Cella As Range
For Each Cella In Intersect(Target.EntireRow.Cells, Target.Parent.UsedRange)
    Debug.Print Cella.Address
    If Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "Descrizione" Then
        Security = Cella.Text
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "ISIN" Then
        ISIN = Cella.Text
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "Valoren" Then
        Valoren = Cella.Text
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "Crncy" Then
        Crncy = Cella.Text
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "Price" Then
        Price = Cella.value
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "I_Pos" Then
        Client = Cella.Parent.Cells(1, Cella.Column)
        Portfolio = Cella.Parent.Cells(2, Cella.Column)
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "O_Pos" Then
        O_Pos = Cella.value
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "O_Amnt" Then
        Amnt = Cella.value
    ElseIf Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Cella.Column) = "O_Ctrlv" Then
        If O_Pos <> 0 Then
            Set O = New clsOrdine
            O.Client = Client
            O.Portfolio = Portfolio
            O.Security = Security
            O.ISIN = ISIN
            O.Crncy = Crncy
            O.Price = CLng(Price)
            O.Number = CLng(O_Pos)
            O.Amount = CLng(Amnt)
            If Len(O.Portfolio) = 8 Then
                O.Bank = "BdS"
            Else
                O.Bank = "UBS"
            End If
            Ordini.Add O
        End If
    End If
Next Cella

If Ordini.Count > 0 Then
    wsName = Left(O.Security, 8) & " (" & ISIN & ")"
    If morfunctions.SheetExists(ActiveWorkbook, wsName) Then
        Set OrdersWs = ActiveWorkbook.Worksheets(wsName)
        Set OrdersTbl = OrdersWs.ListObjects("Ordini")
    Else
        'Crea un file degli ordini
        Set OrdersWs = Worksheets.Add()
        With OrdersWs
            .Name = wsName
            .Cells(1, 1) = Security
            .Cells(2, 1) = "NV:"
            .Cells(2, 2) = Valoren
            .Cells(2, 2).HorizontalAlignment = xlLeft
            .Cells(2, 3) = "ISIN:"
            .Cells(2, 3).HorizontalAlignment = xlRight
            .Cells(2, 4) = ISIN
            .Cells(2, 4).HorizontalAlignment = xlLeft
            .Cells(4, 1) = "Banca"
            .Cells(4, 2) = "Portafoglio"
            .Cells(4, 3) = "Cliente"
            .Cells(4, 4) = "Tipo"
            .Cells(4, 5) = "Quantit" & ChrW(224)
            .Cells(4, 6) = "Imp."
            .Cells(4, 6).HorizontalAlignment = xlRight
            .Cells(4, 7) = "stimato"
            .Cells(4, 8) = "Conto"
            .Cells(4, 9) = "Osservazioni"
            .Cells(4, 10) = "Aggiornato"
            .Cells(4, 11) = "Trasmesso"
            .Range("$A$1:$I$1").Style = "Heading 1"
            .ListObjects.Add(xlSrcRange, Range("$A$4:$K$4"), , xlYes).Name = "Ordini"
            .ListObjects("Ordini").TableStyle = "TableStyleMedium16"
        End With
        Set OrdersTbl = OrdersWs.ListObjects("Ordini")
    End If
    For Each Ordine In Ordini
        Set OrderRw = OrdersTbl.ListRows.Add
        With OrderRw
            .Range(1) = Ordine.Bank
            .Range(2) = Ordine.Portfolio
            .Range(3) = Ordine.Client
            If Ordine.Number > 0 Then .Range(4) = "Ach" Else .Range(4) = "Vte"
            .Range(5) = Ordine.Number
            .Range(6) = Ordine.Crncy
            .Range(7) = Ordine.Amount
            .Range(8) = Ordine.Crncy
            .Range(10) = Now
        End With
    Next Ordine
    
    With OrdersTbl
        'Ordina tabella (prima di eliminare le doppie)
        With .Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=OrdersTbl.DataBodyRange.Columns(1), SortOn:=xlSortOnValues, Order _
                :=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=OrdersTbl.DataBodyRange.Columns(3), SortOn:=xlSortOnValues, Order _
                :=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add2 Key:=OrdersTbl.DataBodyRange.Columns(10), SortOn:=xlSortOnValues, Order _
                :=xlDescending, DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        'Prima di elimanare i doppi, copia eventuali osservazioni nell'ordine pi� recente
        For Each Riga In OrdersTbl.DataBodyRange.Rows
            If Riga.Cells(2) = Riga.Cells(2).Offset(1, 0) And Riga.Cells(9).Offset(1).Text <> "" Then
            'Copia Osservazioni
                Riga.Cells(9).value = Riga.Cells(9).Offset(1).value
            End If
        Next Riga
        
        'Elimina eventuali ordini doppi(il pi� vecchio = Il secondo dopo ordinamento per ora decrescente)
        .Range.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
        'Formata tabella
        .ShowTotals = True
        .ListColumns("Conto").TotalsCalculation = xlTotalsCalculationNone
        .ListColumns("Trasmesso").TotalsCalculation = xlTotalsCalculationNone
        .ListColumns("Quantit" & ChrW(224)).TotalsCalculation = xlTotalsCalculationSum
        .ListColumns("stimato").TotalsCalculation = xlTotalsCalculationSum
        .Range.Rows.RowHeight = 20
        .Range.VerticalAlignment = xlCenter
        .Range.Rows(1).RowHeight = 30
        .Range.Rows(1).VerticalAlignment = xlTop
        .HeaderRowRange.Columns(9).WrapText = True
        .HeaderRowRange.Columns(10).WrapText = True
        .DataBodyRange.Columns(2).HorizontalAlignment = xlLeft
        .Range.Columns(5).HorizontalAlignment = xlRight
        .Range.Columns(5).IndentLevel = 1
        .Range.Columns(5).NumberFormat = "#,##0"
        .Range.Columns(6).HorizontalAlignment = xlRight
        .HeaderRowRange.Columns(7).HorizontalAlignment = xlLeft
        .Range.Columns(7).NumberFormat = "#,##0"
        .DataBodyRange.Columns(8).HorizontalAlignment = xlCenter
        .Range.Columns.AutoFit
        
        'Definitsci colere tab; rosso se vendita / verde se acquisto
        With OrdersWs.Tab
            If OrdersTbl.TotalsRowRange.Cells(1, OrdersTbl.ListColumns("Quantit" & ChrW(224)).Index) > 0 Then
                .Color = 5287936
                .TintAndShade = 0.65
            Else
                .Color = 255
                .TintAndShade = 0.65
            End If
        End With
    End With
End If
fine:
Application.Calculation = Calcolazione
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub
Sub TestAgentRun()
Shell "C:\Program Files\Sequentum\Sequentum Enterprise\RunAgent.exe D:\Cloud\OneDrive\Agents\MS_Scrape_Lists\MS_Scrape_Lists.scg", vbMinimizedNoFocus
End Sub

Sub AggiornadatiSAA()

ActiveWorkbook.Connections("Query - Posizioni").Refresh

''Aria2c must be installed and in the user path
'
'    Selection.Select
'    failed_updates = ""
'    watchlist_path = "B:\Posizioni"
'    portfolio_path = "B:\Portafogli\LUKB"
'
  

    
    
''Sezione disabilitata temporanmenate in quanto produce un errore
'    portfolio_name = "OAU-Trader"
'    Aria2c_CookieFile_Download "https://boersenundmaerkte.lukb.ch/lukb/portfolio/show/export.xls?id=3821494&ds=trader", "D:\Cloud\OneDrive\Cookies\lukb-ch.txt", _
'        portfolio_path, portfolio_name & "_new.xls", "true"
'    If file_first_line(portfolio_path & "\" & portfolio_name & "_new.xls") <> "" Then
'        'Il file non � un HTML con la prima linea vuota
'        If FileExists(portfolio_path & "\" & portfolio_name & ".xls") Then Kill portfolio_path & "\" & portfolio_name & ".xls"
'        Name portfolio_path & "\" & portfolio_name & "_new.xls" As _
'             portfolio_path & "\" & portfolio_name & ".xls"
'    Else ' The first line is empty, as ina a HTML file
'        failed_updates = failed_updates & portfolio_name & "_new.xls " & Chr(10)
'    End If
'
'    portfolio_name = "OAU-AssetClass"
'    Aria2c_CookieFile_Download "https://boersenundmaerkte.lukb.ch/lukb/portfolio/show/export.xls?id=3821494&assetClass=true", "D:\Cloud\OneDrive\Cookies\lukb-ch.txt", _
'        portfolio_path, portfolio_name & "_new.xls", "true"
'    If file_first_line(portfolio_path & "\" & portfolio_name & "_new.xls") <> "" Then
'        'Il file non � un HTML con la prima linea vuota
'        If FileExists(portfolio_path & "\" & portfolio_name & ".xls") Then Kill portfolio_path & "\" & portfolio_name & ".xls"
'        Name portfolio_path & "\" & portfolio_name & "_new.xls" As _
'             portfolio_path & "\" & portfolio_name & ".xls"
'    Else ' The first line is empty, as ina a HTML file
'        failed_updates = failed_updates & portfolio_name & "_new.xls " & Chr(10)
'    End If
'
'        portfolio_name = "OGAB-Trader"
'    Aria2c_CookieFile_Download "https://boersenundmaerkte.lukb.ch/lukb/portfolio/show/export.xls?id=3822628&ds=trader", "D:\Cloud\OneDrive\Cookies\lukb-ch.txt", _
'        portfolio_path, portfolio_name & "_new.xls", "true"
'    If file_first_line(portfolio_path & "\" & portfolio_name & "_new.xls") <> "" Then
'        'Il file non � un HTML con la prima linea vuota
'        If FileExists(portfolio_path & "\" & portfolio_name & ".xls") Then Kill portfolio_path & "\" & portfolio_name & ".xls"
'        Name portfolio_path & "\" & portfolio_name & "_new.xls" As _
'             portfolio_path & "\" & portfolio_name & ".xls"
'    Else ' The first line is empty, as ina a HTML file
'        failed_updates = failed_updates & portfolio_name & "_new.xls " & Chr(10)
'    End If
'
'    portfolio_name = "OGAB-AssetClass"
'    Aria2c_CookieFile_Download "https://boersenundmaerkte.lukb.ch/lukb/portfolio/show/export.xls?id=3822628&assetClass=true", "D:\Cloud\OneDrive\Cookies\lukb-ch.txt", _
'        portfolio_path, portfolio_name & "_new.xls", "true"
'    If file_first_line(portfolio_path & "\" & portfolio_name & "_new.xls") <> "" Then
'        'Il file non � un HTML con la prima linea vuota
'        If FileExists(portfolio_path & "\" & portfolio_name & ".xls") Then Kill portfolio_path & "\" & portfolio_name & ".xls"
'        Name portfolio_path & "\" & portfolio_name & "_new.xls" As _
'             portfolio_path & "\" & portfolio_name & ".xls"
'    Else ' The first line is empty, as ina a HTML file
'        failed_updates = failed_updates & portfolio_name & "_new.xls " & Chr(10)
'    End If
'
'    If failed_updates <> "" Then
'        Message = failed_updates & "presenta(no) un problema." & Chr(10) _
'                & "Non � stato possible scaricare il file o � stato scaricato un file HTML." & Chr(10) _
'                & "I file esistenti non sono stati sovrascritti" & Chr(10) _
'                & "I prezzi sono quindi quelli dell'ultimo aggiornamento riuscito e non attuali"
'        MsgBox Message, vbExclamation, "Errore aggiornamento"
'    End If
'
''Fine sezione disabilitata
    
    
DoEvents
MessaggioInBarraDiStato "Ultimo aggiornamento: " & VBA.Format(Now(), "dd.mm.yy hh:mm:ss"), 0
Exit Sub
End Sub

Private Sub AggiornaOrigine(csvUrl As String, SavePath As String, CsvFileName As String, Optional UserName As String, Optional UserPwd As String, Optional Encoded As Boolean)
If Encoded Then
    ScaricaFileEncoded csvUrl, UserName, UserPwd, SavePath, CsvFileName, True, Encoded
Else
    ScaricaFile csvUrl, UserName, UserPwd, SavePath, CsvFileName, True
End If
DoEvents
End Sub

Private Sub AggiornaQuery(Oggetto As String, Syncronous As Boolean)
' Se syncronous � vero la querai si aggiorna in Background e la macro non attende il suo completamento
With ThisWorkbook.Connections("Query - " & Oggetto)
    On Error GoTo MessaggioErrore
    .OLEDBConnection.BackgroundQuery = Syncronous
    .Refresh
    On Error GoTo 0
End With
MessaggioInBarraDiStato "Ultimo aggiornamento query: " & VBA.Format(Now(), "dd.mm.yy hh:mm:ss"), 0
Exit Sub
MessaggioErrore:
MsgBox "Purtroppo non � stato possibile aggiornare la query " & Oggetto & ". Verfica che i dati di origine siano disponibili.", vbCritical, "Errore aggiornamento query"
End Sub

Sub ResetLukbId(ByRef AreaSelezionata As Range)
Dim riga_intestazioni As Range, colonna_smo_id_origine As Range, riga_search_id_posizioni As Range, colonna_lukb_id_posizioni As Range
Dim lukb_to_reset As Range
Dim search_id As Variant
Dim prima_riga_selezione As Range
Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean

On Error GoTo fine
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False

Set prima_riga_selezione = AreaSelezionata.Cells(1, 1).EntireRow
Dim rng_colonna_smo_id_origine As Range
Set rng_colonna_smo_id_origine = AreaSelezionata.Parent.Cells.Find(What:="smo_id", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_colonna_smo_id_origine Is Nothing Then
Set colonna_smo_id_origine = rng_colonna_smo_id_origine.EntireColumn
End If

search_id = Intersect(prima_riga_selezione, colonna_smo_id_origine)

Dim rng_riga_search As Range
Set rng_riga_search = Posizioni.Cells.Find(What:=search_id, After:=Cells(1, 1), LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
If Not rng_riga_search Is Nothing Then Set riga_search_id_posizioni = rng_riga_search.EntireRow

Dim rng_col_lukb As Range
Set rng_col_lukb = Posizioni.Cells.Find(What:="LUKB_id", After:=Cells(1, 1), LookIn:=xlValues _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
If Not rng_col_lukb Is Nothing Then Set colonna_lukb_id_posizioni = rng_col_lukb.EntireColumn

If riga_search_id_posizioni Is Nothing Or colonna_lukb_id_posizioni Is Nothing Then GoTo fine
Set lukb_to_reset = Intersect(riga_search_id_posizioni, colonna_lukb_id_posizioni)
lukb_to_reset.value = ""
ActiveWorkbook.Connections("Query - Posizioni").Refresh

fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating

End Sub

Sub PorfolioSheetSelectionChange(ByRef Target As Range)
Dim sec_id_text As String
Dim sec_id_elements() As String

On Error GoTo Overflow
If Target.Cells.Count = 1 Then
    SAAProc.ValoreSelezione = Target.value
    sec_id_text = Intersect(Target.Parent.Range("col_ID"), Target.EntireRow)
    sec_id_elements = Split(sec_id_text, ":")
    If UBound(sec_id_elements) >= 1 Then
        SAAProc.sec_id = Join(Array(sec_id_elements(0), sec_id_elements(1)), ":")
    End If
End If
'Triggered when too many cells selected; i.e. All Cells
Overflow:
End Sub

Sub PortfolioSheetChange(ByRef Target As Range)
Dim row_id_text As String
Dim row_id_elements() As String

If Target.Cells.Count = 1 And Target.Cells(1, 1).HasFormula = False Then
    'Register current setting and apply manual calculation
    'SAAProc.calculation_setting = Application.Calculation
    'Application.Calculation = xlCalculationManual
    
    On Error GoTo fine
    
    row_id_text = Intersect(Target.Parent.Range("col_ID"), Target.EntireRow).value
    row_id_elements = Split(row_id_text, ":")
    If UBound(row_id_elements) >= 1 Then
        SAAProc.row_id = Join(Array(row_id_elements(0), row_id_elements(1)), ":")
    End If

    
    If SAAProc.row_id <> SAAProc.sec_id Then
        SAAProc.ResetPosizioni Target
        MORProcedures.Aggiorna_Hyperlinks Target.Parent.Rows(Target.Row)
        'Aggiorna_Campi_MarketScreener Target
    End If
    
    If Target.Parent.Cells(Target.Parent.Range("rwIntestazioni").Row, Target.Column) = "I_Pos" Then
        'XXX
        With Valori
            Set Valore = .Cells.Find(Target.Parent.Cells(Target.Row, Target.Parent.Range("col_ID").Column))
            Set Posizione = .Cells(Valore.Row, .Cells.Find("Position").Column)
            Posizione = Posizione + Target - ValoreSelezione
            ValoreSelezione = Target.value
        End With
    End If
fine:
    'Application.Calculation = SAAProc.calculation_setting
End If
End Sub

Sub AggiornaPosizione(ByRef rgOrdini As Range)

Dim Ordine As Double, VecchiaPosizione As Double, NuovaPosizione As Double, VecchioSaldo As Double, NuovoSaldo As Double, Conversione As Double, PctOrdine As Double, PctLiq As Double, PctBnd As Double, PctEqu As Double, PctImm As Double, PctCom As Double, PctHfd As Double
Dim Formula As String, VecchioCommento As String, NuovoCommento As String, ContoAlternativo As String
Dim ColonnaCrncy As Range, ColonnaPrice As Range, ColonnaFxR As Range, ColonnaPos As Range, ColonnaAmnt As Range, ColonnaCtrlv As Range, ColonnaPosPct As Range, ColonnaOrdPct As Range, _
    RigaIntestazioni As Range, RigaClient As Range, RigaMRif As Range, RigaCC As Range, RigaAC As Range, RigaSC As Range, RigaLiq As Range, RigaBnd As Range, RigaEqu As Range, RigaImm As Range, RigaCom As Range, RigaHfd As Range, _
    CellaMRif As Range, CellaCC As Range, CellaDescrizione As Range, CellaCrncy As Range, CellaPrice As Range, CellaPosizione As Range, CellaOrdine As Range, CellaAmnt As Range, CellaCtrlv As Range, CellaPosPct As Range, CellaOrdPct As Range, Temp As Variant

' Order log Object
Dim LogTbl As ListObject
Dim LogRw As ListRow
 
Dim ScartoRighe, ScartoColonne, NoColOrdine As Integer

On Error GoTo fine
Dim prevEvents As Boolean
Dim prevScreenUpdating As Boolean
prevEvents = Application.EnableEvents
prevScreenUpdating = Application.ScreenUpdating
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.ScreenUpdating = False

Set RigaIntestazioni = rgOrdini.Parent.Range("rwIntestazioni")
NoRigaIntestazioni = RigaIntestazioni.Row
NoColOrdine = rgOrdini.Cells(1, 1).Column

Dim CellaOrdine As Range
For Each CellaOrdine In rgOrdini.Columns(1).Cells
    If Intersect(CellaOrdine.EntireColumn, RigaIntestazioni) = "O_Pos" And CellaOrdine <> "" And CellaOrdine <> 0 Then
        'Raccogli i dati dell'ordine
        Ordine = CellaOrdine.value
        
        Dim rng_ColonnaAC As Range
Set rng_ColonnaAC = CellaOrdine.Parent.Cells.Find(What:="AC", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaAC Is Nothing Then
Set ColonnaAC = rng_ColonnaAC.EntireColumn
End If

        Dim rng_ColonnaSC As Range
Set rng_ColonnaSC = CellaOrdine.Parent.Cells.Find(What:="SC", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaSC Is Nothing Then
Set ColonnaSC = rng_ColonnaSC.EntireColumn
End If

        Dim rng_ColonnaDescrizione As Range
Set rng_ColonnaDescrizione = CellaOrdine.Parent.Cells.Find(What:="Descrizione", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaDescrizione Is Nothing Then
Set ColonnaDescrizione = rng_ColonnaDescrizione.EntireColumn
End If

        Dim rng_ColonnaCrncy As Range
Set rng_ColonnaCrncy = CellaOrdine.Parent.Cells.Find(What:="Crncy", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaCrncy Is Nothing Then
Set ColonnaCrncy = rng_ColonnaCrncy.EntireColumn
End If

        Dim rng_ColonnaFxR As Range
Set rng_ColonnaFxR = CellaOrdine.Parent.Cells.Find(What:="FxR", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaFxR Is Nothing Then
Set ColonnaFxR = rng_ColonnaFxR.EntireColumn
End If

        Dim rng_ColonnaPrice As Range
Set rng_ColonnaPrice = CellaOrdine.Parent.Cells.Find(What:="Price", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaPrice Is Nothing Then
Set ColonnaPrice = rng_ColonnaPrice.EntireColumn
End If


        Dim rng_ColonnaPos As Range
Set rng_ColonnaPos = CellaOrdine.Parent.Cells.Find(What:="I_Pos", After:=Cells(NoRigaIntestazioni, NoColOrdine), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, SearchFormat:=False)
If Not rng_ColonnaPos Is Nothing Then
Set ColonnaPos = rng_ColonnaPos.EntireColumn
End If

        Dim rng_ColonnaPam As Range
Set rng_ColonnaPam = CellaOrdine.Parent.Cells.Find(What:="I_Price", After:=Cells(NoRigaIntestazioni, NoColOrdine), LookIn:=xlFormulas2, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, SearchFormat:=False)
If Not rng_ColonnaPam Is Nothing Then
Set ColonnaPam = rng_ColonnaPam.EntireColumn
End If

        Dim rng_ColonnaPosPct As Range
Set rng_ColonnaPosPct = CellaOrdine.Parent.Cells.Find(What:="I_Pct", After:=Cells(NoRigaIntestazioni, NoColOrdine), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, SearchFormat:=False)
If Not rng_ColonnaPosPct Is Nothing Then
Set ColonnaPosPct = rng_ColonnaPosPct.EntireColumn
End If

        Dim rng_ColonnaAmnt As Range
Set rng_ColonnaAmnt = CellaOrdine.Parent.Cells.Find(What:="O_Amnt", After:=Cells(NoRigaIntestazioni, NoColOrdine), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaAmnt Is Nothing Then
Set ColonnaAmnt = rng_ColonnaAmnt.EntireColumn
End If

        Dim rng_ColonnaCtrlv As Range
Set rng_ColonnaCtrlv = CellaOrdine.Parent.Cells.Find(What:="O_Ctrlv", After:=Cells(NoRigaIntestazioni, NoColOrdine), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaCtrlv Is Nothing Then
Set ColonnaCtrlv = rng_ColonnaCtrlv.EntireColumn
End If

        Dim rng_ColonnaOrdPct As Range
Set rng_ColonnaOrdPct = CellaOrdine.Parent.Cells.Find(What:="O_Pct", After:=Cells(NoRigaIntestazioni, NoColOrdine), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
If Not rng_ColonnaOrdPct Is Nothing Then
Set ColonnaOrdPct = rng_ColonnaOrdPct.EntireColumn
End If


        
        Set CellaAC = Intersect(CellaOrdine.EntireRow, ColonnaAC).End(xlUp)
        Set CellaSC = Intersect(CellaOrdine.EntireRow, ColonnaSC).End(xlUp)
        Set CellaDescrizione = Intersect(CellaOrdine.EntireRow, ColonnaDescrizione)
        Set CellaCrncy = Intersect(CellaOrdine.EntireRow, ColonnaCrncy)
        Set CellaPosizione = CellaOrdine.Parent.Cells(CellaOrdine.Row, ColonnaPos.Column)
        Set CellaPam = Intersect(CellaOrdine.EntireRow, ColonnaPam)
        Set CellaPosPct = CellaOrdine.Parent.Cells(CellaOrdine.Row, ColonnaPosPct.Column)
        Set CellaFxR = CellaOrdine.Parent.Cells(CellaOrdine.Row, ColonnaFxR.Column)
        Set CellaPrice = CellaOrdine.Parent.Cells(CellaOrdine.Row, ColonnaPrice.Column)
        Set CellaAmnt = Intersect(CellaOrdine.EntireRow, ColonnaAmnt)
        Set CellaCtrlv = Intersect(CellaOrdine.EntireRow, ColonnaCtrlv)
        Set CellaOrdPct = Intersect(CellaOrdine.EntireRow, ColonnaOrdPct)
        
        'Se non c'� il conto corrente
        '& CellaCrncy.Value
        If ColonnaDescrizione.Cells.Find(What:="CC " & CellaCrncy.value, LookIn:=xlFormulas2, LookAt:=xlWhole) Is Nothing Then
            ContoAlternativo = InputBox("Su quale CC desideri contabilizzare la transazione?", _
                "Non c'� un conto corrente nella moneta della transazione", CellaOrdine.Parent.Range("MRif"))
            Set RigaCC = ColonnaDescrizione.Find(What:="CC " & ContoAlternativo, LookIn:=xlFormulas2, LookAt:=xlWhole).EntireRow
        Else
            ContoAlternativo = ""
            Set RigaCC = ColonnaDescrizione.Find(What:="CC " & CellaCrncy.value, LookIn:=xlFormulas2, LookAt:=xlWhole).EntireRow

        End If
        
        'Determina il conto liquidita e la MRif rilevante
        Set CellaCC = Intersect(RigaCC, ColonnaPos)
        Set RigaMRif = ColonnaDescrizione.Find("CC " & CellaOrdine.Parent.Range("MRif")).EntireRow
        Set CellaMRif = Intersect(RigaMRif, ColonnaPos)
        VecchiaPosizioneFormula = CellaPosizione.Formula
        VecchiaPosizione = CellaPosizione.value
        VecchioSaldo = CellaCC.value

        If CellaCC.Row <> CellaPosizione.Row Then ' non � una modifica del saldo valutario
            'Se la transazione � contabilizzata su un conto alternativo
            If ContoAlternativo <> "" Then
                Conversione = Intersect(CellaOrdine.EntireRow, ColonnaFxR) / Intersect(RigaCC, ColonnaFxR)
                CellaCC.value = VecchioSaldo - (CellaAmnt.value * Conversione)
            Else
                CellaCC.value = VecchioSaldo - CellaAmnt '.value
            End If
        Else ' � uma modifica del saldo, dunque addebita/accredita moneta di riferimento
            CellaMRif.value = CellaMRif.value - CellaCtrlv.value
        End If
        
        ' Aggiorna la posizione
        If Left(VecchiaPosizioneFormula, 1) <> "=" Then
            VecchiaPosizioneFormula = "=" & VecchiaPosizione
         End If
        CellaPosizione.Formula = VecchiaPosizioneFormula & "+" & Ordine
        If CellaPosizione.value = 0 Then
            If Not CellaPosizione.Comment Is Nothing Then CellaPosizione.ClearComments
        Else
            If CellaPosizione.Comment Is Nothing Then
                CellaPosizione.AddComment
                VecchioCommento = ""
            Else
                VecchioCommento = CellaPosizione.Comment.Text & Chr(10)
            End If
            NuovoCommento = Ordine & " @ " & rgOrdini.Parent.Cells(CellaPosizione.Row, rgOrdini.Parent.Range("clPrice").Column).Text & " al " & Format(Now(), "dd.mm.yy hh:mm")
            With CellaPosizione.Comment
                .Text VecchioCommento & NuovoCommento
                .Shape.TextFrame.Characters.Font.Name = "Calibri"
                .Shape.TextFrame.Characters.Font.Size = 8
            End With
        End If
        
        'Aggiorna il prezzo medio
        If VecchiaPosizioneFormula = "=" Or VecchiaPosizione = 0 Then
            NuovoPam = CellaPrice.value
        Else
            If Ordine > 0 And CellaPam.value > 0 Then
                NuovoPam = ((VecchiaPosizione * CellaPam.value) + (Ordine * CellaPrice.value)) / CellaPosizione.value
            Else
                If CellaPosizione.value = 0 Then
                    NuovoPam = ""
                Else
                    NuovoPam = CellaPam.value
                End If
            End If
        End If
        CellaPam.Formula = NuovoPam
    
        'Leggi le percentuali TAA successive all'ordine
        Set RigaClient = CellaOrdine.Parent.Cells.Find(What:="Intestato a", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).EntireRow
        On Error Resume Next ' in case one of the essets classes is not present, an error will be triggered
            Set RigaLiq = CellaOrdine.Parent.Cells.Find(What:="Liquidit" & ChrW(224), After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).EntireRow
            Set RigaBnd = CellaOrdine.Parent.Cells.Find(What:="Obbligazioni", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).EntireRow
            Set RigaEqu = CellaOrdine.Parent.Cells.Find(What:="Azioni", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).EntireRow
            Set RigaImm = CellaOrdine.Parent.Cells.Find(What:="Immobiliare", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).EntireRow
            Set RigaCom = CellaOrdine.Parent.Cells.Find(What:="Materia Prime", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).EntireRow
            Set RigaHfd = CellaOrdine.Parent.Cells.Find(What:="Altro", After:=Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).EntireRow
            PctLiq = Intersect(RigaLiq, ColonnaPosPct)
            PctBnd = Intersect(RigaBnd, ColonnaPosPct)
            PctEqu = Intersect(RigaEqu, ColonnaPosPct)
            PctImm = Intersect(RigaImm, ColonnaPosPct)
            PctCom = Intersect(RigaCom, ColonnaPosPct)
            PctHfd = Intersect(RigaHfd, ColonnaPosPct)
        On Error GoTo fine

        'Esporta i dati nella tabella log ordini se valore non Nothing
        Set LogTbl = RegistroOrdini.ListObjects("Registro_Ordini")
        Set LogRw = LogTbl.ListRows.Add
        With LogRw
            .Range(1) = Now() 'Format(Now(), "dd.mm.yyyy")
            .Range(2) = Now() 'Format(Now(), "hh:mm")
            .Range(3) = Intersect(RigaClient, ColonnaPos)
            .Range(4) = CellaAC.value
            .Range(5) = CellaSC.value
            .Range(6) = CellaDescrizione.value
            If Ordine < 0 Then .Range(7) = "Sell" Else .Range(7) = "Buy"
            .Range(8) = Ordine
            .Range(9) = CellaPrice.value
            .Range(10) = CellaCrncy.value
            .Range(11).FormulaR1C1 = "=[@Numero]*[@Prezzo]" ' = CellaAmnt.value"
            .Range(12) = CellaFxR.value
            .Range(13).FormulaR1C1 = "=[@Importo]*[@Cambio]" '=CellaCtrlv.value
            .Range(16) = CellaOrdPct.value
            .Range(17) = PctLiq
            .Range(18) = PctBnd
            .Range(19) = PctEqu
            .Range(20) = PctImm
            .Range(21) = PctCom
            .Range(22) = PctHfd

        End With

        'Elimina l'ordine registrato
        CellaOrdine.ClearContents
    End If
Next CellaOrdine
fine:
Application.EnableEvents = prevEvents
Application.ScreenUpdating = prevScreenUpdating
End Sub
Sub SetColumnWidthByHeader( _
        ByVal Pct_Column_Width As Double, _
        ByVal Currency_Column_Width As Double)

    Dim rng As Range
    Dim c As Range
    Dim firstColNum As Long
    Dim txt As String
    Dim foundFxR As Range
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject

    Set ws = ActiveSheet

    '---------------------------------------------------------
    ' PRE-RUN: Stop background queries and freeze UI
    '---------------------------------------------------------
    On Error Resume Next

    ' Stop QueryTables
    For Each qt In ws.QueryTables
        qt.CancelRefresh
    Next qt

    ' Stop PowerQuery ListObjects
    For Each lo In ws.ListObjects
        lo.QueryTable.CancelRefresh
    Next lo

    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    '---------------------------------------------------------
    ' MAIN LOGIC
    '---------------------------------------------------------

    ' Resolve the named range
    On Error GoTo ErrHandler
    Set rng = Range("rwIntestazioni")
    On Error GoTo 0

    ' Find the column containing "FxR"
    Set foundFxR = Nothing
    For Each c In rng.Cells
        If Trim(CStr(c.value)) = "FxR" Then
            Set foundFxR = c
            Exit For
        End If
    Next c

    If foundFxR Is Nothing Then
        MsgBox "Header 'FxR' not found in rwIntestazioni.", vbCritical
        GoTo Cleanup
    End If

    firstColNum = foundFxR.Column

    ' Loop through each cell in the header row
    For Each c In rng.Cells
        If c.Column > firstColNum Then

            txt = Trim(CStr(c.value))

            ' Empty cells remain untouched
            If Len(txt) = 0 Then GoTo NextCell

            ' Hide column if header contains "Visible"
            If StrComp(txt, "Visible", vbTextCompare) = 0 Then
                c.EntireColumn.Hidden = True
                GoTo NextCell
            End If

            ' Ensure visible before applying width rules
            c.EntireColumn.Hidden = False

            ' Percent columns: contains "%" OR "_Pct"
            If InStr(1, txt, "%", vbTextCompare) > 0 _
               Or InStr(1, txt, "_Pct", vbTextCompare) > 0 Then

                c.EntireColumn.ColumnWidth = Pct_Column_Width

            ' Currency columns: contains "_" (but not _Pct)
            ElseIf InStr(1, txt, "_", vbTextCompare) > 0 Then
                c.EntireColumn.ColumnWidth = Currency_Column_Width
            End If

        End If

NextCell:
    Next c

    GoTo Cleanup

'---------------------------------------------------------
' ERROR HANDLING
'---------------------------------------------------------
ErrHandler:
    MsgBox "Named range 'rwIntestazioni' not found.", vbCritical

'---------------------------------------------------------
' CLEANUP: Restore Excel state
'---------------------------------------------------------
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc

End Sub


Sub ApplyHeaderFormatting()
    Call SetColumnWidthByHeader(6.1, 8.1)
End Sub

Sub SetWorkbookDefaultFont( _
        ByVal FontName As String, _
        ByVal FontSize As Double)

    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject

    '---------------------------------------------------------
    ' PRE-RUN: Stop background queries and freeze UI
    '---------------------------------------------------------
    On Error Resume Next

    ' Stop QueryTables
    For Each ws In ActiveWorkbook.Worksheets
        For Each qt In ws.QueryTables
            qt.CancelRefresh
        Next qt
    Next ws

    ' Stop PowerQuery ListObjects
    For Each ws In ActiveWorkbook.Worksheets
        For Each lo In ws.ListObjects
            lo.QueryTable.CancelRefresh
        Next lo
    Next ws

    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual

    '---------------------------------------------------------
    ' MAIN LOGIC: Apply font to all worksheets
    '---------------------------------------------------------
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.Font.Name = FontName
        ws.Cells.Font.Size = FontSize
    Next ws

    '---------------------------------------------------------
    ' CLEANUP: Restore Excel state
    '---------------------------------------------------------
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc

End Sub

Sub ApplyWorkbookDefaultFont()
    Call SetWorkbookDefaultFont("Aptos", 9)
End Sub






