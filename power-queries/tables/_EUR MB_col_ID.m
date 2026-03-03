// Query   : #"'EUR MB'!col_ID"
// Category: tables
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-03-03T17:39:11+00:00

let
        Origine = Excel.CurrentWorkbook(){[Name = "'EUR MB'!col_ID"]}[Content],
        #"Filtrate righe" = Table.SelectRows(Origine, each Text.Contains([Column1], ":"))
    in
        #"Filtrate righe"
