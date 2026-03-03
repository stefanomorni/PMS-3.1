// Query   : Porfolio_Positions
// Category: tables
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-03-03T17:39:11+00:00

let
        Origine = Table.Combine(
            {#"'CHF M'!col_ID", #"'CHF MA'!col_ID", #"'EUR MB'!col_ID", #"'EUR M'!col_ID", #"'AGCO'col_ID"}
        ),
        #"Rimossi duplicati" = Table.Distinct(Origine),
        #"Rinominate colonne" = Table.RenameColumns(#"Rimossi duplicati", {{"Column1", "smo_id"}})
    in
        #"Rinominate colonne"
