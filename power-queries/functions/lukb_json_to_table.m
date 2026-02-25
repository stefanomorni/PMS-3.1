// Query   : lukb_json_to_table
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

(json_data as record, listing_ids as list, fields_codes as list) =>
    let
        field_codes_elements_list = Table.FromList(
            fields_codes, Splitter.SplitTextByDelimiter(":"), null, null, ExtraValues.Error
        ),
        fields_root_tbl = Table.RemoveColumns(field_codes_elements_list, {"Column3", "Column4"}),
        fields_root_list = fields_root_tbl[Column1],
        // corrsponsing to level 2 nodes of Json: Field, without type specifier
        fields_type_table = Table.CombineColumns(
            fields_root_tbl, {"Column1", "Column2"}, Combiner.CombineTextByDelimiter(":", QuoteStyle.None),
            "Field_ids"
        ),
        fields_type_list = fields_type_table[Field_ids],
        data_table = Record.ToTable(json_data),
        expanded_records_tbl = fn_expand_all_records(data_table, null, null, null, ":"),
        // reorder the columns as the original field_type_list
        fields_type_ex_value_prefix_list = List.Transform(
            Table.ColumnNames(expanded_records_tbl), each {_, Text.Replace(_, "Value:", "")}
        ),
        trimmed_headers_tbl = Table.RenameColumns(expanded_records_tbl, fields_type_ex_value_prefix_list),
        reorder_columns_by_field_types = Table.ReorderColumns(
            trimmed_headers_tbl, fields_type_list, MissingField.UseNull
        ),
        // remove the ripetitive ":value" from the column headings
        fields_type_ex_value_suffix_list = List.Transform(
            Table.ColumnNames(reorder_columns_by_field_types), each {_, Text.Replace(_, ":value", "")}
        ),
        renamed_data_table = Table.RenameColumns(reorder_columns_by_field_types, fields_type_ex_value_suffix_list),
        // convert the column types according to the first rows to ensure numbers are recognised as such for calculation
        convert_column_types = fn_convert_column_types(renamed_data_table),
        // Add calculated values Calculated value Spread %
        calc_bid_ask_spread = Table.AddColumn(
            convert_column_types, "Bid-Ask Spread %", each try ([ASK] - [BID]) / ([BID] + [ASK] / 2) otherwise null
        ),
        // Add index column based on the position in listing_ids
        add_sorting_column = Table.AddColumn(
            calc_bid_ask_spread, "SortOrder", (row) => List.PositionOf(listing_ids, row[Name])
        ),
        // Sort by the new index column
        sort_rows_by_listing_ids = Table.Sort(add_sorting_column, {{"SortOrder", Order.Ascending}}),
        remove_temp_columns = Table.RemoveColumns(sort_rows_by_listing_ids, {"SortOrder"}),
        rename_name_to_listing_id = Table.RenameColumns(remove_temp_columns, {{"Name", "listing_id"}})
    in
        rename_name_to_listing_id
