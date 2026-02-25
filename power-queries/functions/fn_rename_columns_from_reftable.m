// Query   : fn_rename_columns_from_reftable
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Renamed_Columns_Tbl = (InputTable as table, RefTable as table) as table =>
            let
                Input_Table = InputTable,
                Input_Col_Names_List = Table.ColumnNames(Input_Table),
                Col_Names_List_To_Table = Table.FromList(
                    Input_Col_Names_List, Splitter.SplitByNothing(), null, null, ExtraValues.Error
                ),
                Merge_From_To_Conversion_Table = Table.NestedJoin(
                    Col_Names_List_To_Table, {"Column1"}, RefTable, {"From"}, "RefTable", JoinKind.Inner
                ),
                Expand_Conversion_Table = Table.ExpandTableColumn(
                    Merge_From_To_Conversion_Table, "RefTable", {"To"}, {"To"}
                ),
                Merge_Columns = Table.CombineColumns(
                    Expand_Conversion_Table,
                    {"Column1", "To"},
                    Combiner.CombineTextByDelimiter(",", QuoteStyle.None),
                    "Merged"
                ),
                Add_Mapping_Column = Table.AddColumn(Merge_Columns, "Mapping", each Text.Split([Merged], ",")),
                Mapping_Column_To_List = Add_Mapping_Column[Mapping],
                Apply_Mapping_Ton_Input_Tbl = Table.RenameColumns(Input_Table, Mapping_Column_To_List)
            in
                Apply_Mapping_Ton_Input_Tbl
    in
        Renamed_Columns_Tbl
