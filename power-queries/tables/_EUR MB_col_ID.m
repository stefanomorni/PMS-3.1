// Query   : #"'EUR MB'!col_ID"
// Category: tables
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

let
    Origine = Excel.CurrentWorkbook(){[Name="'EUR MB'!col_ID"]}[Content],
    #"Filtrate righe" = Table.SelectRows(Origine, each Text.Contains([Column1], ":"))
in
    #"Filtrate righe"
