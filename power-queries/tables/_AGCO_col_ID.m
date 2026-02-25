// Query   : #"'AGCO'col_ID"
// Category: tables
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Origine = Excel.CurrentWorkbook(){[Name = "'AGCO'!col_ID"]}[Content],
        #"Filtrate righe" = Table.SelectRows(Origine, each Text.Contains([Column1], ":"))
    in
        #"Filtrate righe"
