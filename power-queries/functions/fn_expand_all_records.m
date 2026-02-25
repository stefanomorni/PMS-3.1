// Query   : fn_expand_all_records
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Source = (
            TableToExpand as table,
            optional ColName as text,
            optional ColNumber as number,
            optional AppendOriginal as logical,
            optional Separator as text
        ) =>
            let
                //If the column number is missing, make it 0
                ColumnNumber = if (ColNumber = null) then 0 else ColNumber,
                //Supplying a ColName parameter overrides the column-finding logic
                ColumnName = if (ColName <> null) then ColName else
                //Find the column name relating to the column number
                Table.ColumnNames(TableToExpand){ColumnNumber},
                //Get a list containing all of the values in the column
                ColumnContents = Table.Column(TableToExpand, ColumnName),
                //Iterate over each value in the column and then
                //If the value is of type table get a list of all of the columns in the table
                //Then get a distinct list of all of these column names
                ColumnsToExpand = List.Distinct(
                    List.Combine(List.Transform(ColumnContents, each if _ is record then Record.FieldNames(_) else {}))
                ),
                //Check if append original is desired, seta it as true if not set
                AppendOriginal = if (AppendOriginal <> null) then AppendOriginal else true,
                //Check if a specific separator is supplied, otherwise use "."
                Separator = if (Separator <> null) then Separator else ".",
                //Append the original column name to the front of each of these column names
                NewColumnNames =
                    if AppendOriginal then
                        List.Transform(ColumnsToExpand, each ColumnName & Separator & _)
                    else
                        List.Transform(ColumnsToExpand, each _),
                //Is there anything to expand in this column?
                CanExpandCurrentColumn = List.Count(ColumnsToExpand) > 0,
                //If this column can be expanded, then expand it
                ExpandedTable =
                    if CanExpandCurrentColumn then
                        Table.ExpandRecordColumn(TableToExpand, ColumnName, ColumnsToExpand, NewColumnNames)
                    else
                        TableToExpand,
                //If the column has been expanded then keep the column number the same, otherwise add one to it
                NextColumnNumber = if CanExpandCurrentColumn then ColumnNumber else ColumnNumber + 1,
                //If the column number is now greater than the number of columns in the table
                //Then return the table as it is
                //Else call the ExpandAll function recursively with the expanded table
                OutputTable =
                    if NextColumnNumber > (Table.ColumnCount(ExpandedTable) - 1) then
                        ExpandedTable
                    else
                        fn_expand_all_records(ExpandedTable, null, NextColumnNumber, AppendOriginal, Separator)
            in
                OutputTable
    in
        Source
