import os


def fix_form():
    path = r"D:\Cloud\OneDrive\Office Junctions\%APPDATA%-Microsoft\AddIns\MORProcedures\vba-files\FrmFinestraInformativa.frm"

    # The correct header for this form
    header = [
        "VERSION 5.00",
        "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmFinestraInformativa ",
        "   ClientHeight    =   1785",
        "   ClientLeft      =   30",
        "   ClientTop       =   330",
        "   ClientWidth     =   7545",
        '   OleObjectBlob   =   "FrmFinestraInformativa.frx":0000',
        "   StartUpPosition =   1  'CenterOwner",
        "End",
        'Attribute VB_Name = "FrmFinestraInformativa"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = False",
        "",
        "",
        "'API function to enable/disable the Excel Window",
        'Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr',
        'Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal bEnable As Long) As Long',
        "",
    ]

    # The code after the API declarations (starting from Dim mlHWnd...)
    code_after = [
        "Dim mlHWnd As LongPtr, mbModal As Boolean, mbDragDrop As Boolean",
        "",
        "Private Sub UserForm_Activate()",
        "",
        "    On Error Resume Next",
        "",
        "    'Find the Excel main window",
        '    mlHWnd = FindWindowA("XLMAIN", Application.Caption)',
        "    mbDragDrop = Application.CellDragAndDrop     'Memorize the current state",
        "",
        "    If CbxModeless.Value Then",
        "        EnableWindow mlHWnd, 1                   'Enable the Window - makes the userform modeless",
        "        'Disable Cell drag/drop, as it causes Excel 97 to GPF",
        "        Application.CellDragAndDrop = False",
        "    Else",
        "        EnableWindow mlHWnd, 0                   'Disable the Window - makes the userform modal",
        "    End If",
        "End Sub",
        "",
        "Private Sub CmdOK_Click()",
        "    Application.CellDragAndDrop = mbDragDrop",
        "    Unload FrmFinestraInformativa",
        "End Sub",
        "",
    ]

    full_content = "\r\n".join(header + code_after)

    # Write as UTF-8 WITHOUT BOM
    with open(path, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(full_content)
    print(f"Fixed Form: {path}")


fix_form()
