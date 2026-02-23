// Query   : lukb_currencies_tbl
// Category: lookups
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

let
    Source = Excel.Workbook(File.Contents("D:\Cloud\OneDrive\MC\InvData\Reference Data.xlsm"), null, true),
    LUKB_Markets_Table = Source{[Item="LUKB_Currencies",Kind="Table"]}[Data],
    #"Renamed Columns" = Table.RenameColumns(LUKB_Markets_Table,{{"value", "cur_code"}, {"description", "cur_des"}, {"id", "cur_id"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns",{{"cur_code", type text}, {"cur_des", type text}, {"cur_id", type text}})
in
    #"Changed Type"
