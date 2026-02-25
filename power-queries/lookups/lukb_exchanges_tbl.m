// Query   : lukb_exchanges_tbl
// Category: lookups
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Source = Excel.Workbook(File.Contents("D:\Cloud\OneDrive\MC\InvData\Reference Data.xlsm"), null, true),
        LUKB_Markets_Table = Source{[Item = "LUKB_Markets", Kind = "Table"]}[Data],
        #"Renamed Columns" = Table.RenameColumns(
            LUKB_Markets_Table, {{"value", "exch_code"}, {"description", "exch_des"}, {"id", "exch_id"}}
        ),
        #"Changed Type" = Table.TransformColumnTypes(
            #"Renamed Columns", {{"exch_id", type text}, {"exch_des", type text}, {"exch_code", type text}}
        )
    in
        #"Changed Type"
