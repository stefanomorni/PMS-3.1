// Query   : fn_range_to_list
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

(
    current_wb_range_name as text,
    optional column_name as nullable text,
    optional row_first_column_value as nullable text
) =>
    let
        // Determine the range type based on input parameters
        range_type =
            if column_name = null and row_first_column_value = null then
                "named_range"
            else if column_name = null then
                if row_first_column_value = "Headers" then
                    "table_headers"
                else
                    "table_row"
            else if row_first_column_value = null then
                "table_column"
            else
                "cell",
        // Get the Excel table from the current workbook
        excel_table = Excel.CurrentWorkbook(){[Name = current_wb_range_name]}[Content],
        // Convert the range to a list based on the determined range type
        range_values_list =
            if range_type = "named_range" then
                if Table.ColumnCount(excel_table) > Table.RowCount(excel_table) then
                    Record.ToList(excel_table{0})
                else
                    excel_table[Column1]
            else if range_type = "table_column" then
                Table.Skip(
                    Table.TransformColumnTypes(
                        Table.DemoteHeaders(Table.SelectColumns(excel_table, {column_name})), {{"Column1", type any}}
                    ),
                    1
                )[Column1]
            else if range_type = "table_headers" then
                Record.ToList(Table.DemoteHeaders(excel_table){0})
            else if range_type = "table_row" then
                let
                    first_column = Table.ColumnNames(excel_table){0},
                    row_index = Table.PositionOf(
                        Table.SelectColumns(excel_table, {first_column}),
                        Record.FromList({row_first_column_value}, {first_column})
                    )
                in
                    if row_index = -1 then
                        null
                    else
                        Record.ToList(excel_table{row_index})
            else if range_type = "cell" then
                let
                    row_index = Table.PositionOf(
                        Table.DemoteHeaders(Table.SelectColumns(excel_table, Table.ColumnNames(excel_table){0})),
                        [
                            Column1 = row_first_column_value
                        ]
                    ) - 1,
                    table_column = Table.Skip(
                        Table.TransformColumnTypes(
                            Table.DemoteHeaders(Table.SelectColumns(excel_table, {column_name})),
                            {{"Column1", type any}}
                        ),
                        1
                    )[Column1]
                in
                    if row_index = -1 then
                        null
                    else
                        {table_column{row_index}}
            else
                null
    in
        range_values_list
