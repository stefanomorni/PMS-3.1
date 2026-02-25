// Query   : fn_convert_column_types
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Origine = (table as table, optional culture as nullable text) as table =>
            let
                InvalidTypes = {type list, type record, type table, type function, type none, type type},
                Top100Rows = Table.FirstN(table, 100),
                //we use up to 100 rows to establish a column type
                ColumnNameList = Table.ColumnNames(Top100Rows),
                ColumnDataLists = List.Accumulate(
                    ColumnNameList, {}, (accumulated, i) => accumulated & {Table.Column(Top100Rows, i)}
                ),
                ColumnTypes = List.Transform(ColumnDataLists, (i) => List.ItemType(i)),
                TransformList = List.Select(
                    List.Zip({ColumnNameList, ColumnTypes}),
                    (i) => not List.AnyTrue(List.Transform(InvalidTypes, (j) => Type.Is(i{1}, j)))
                ),
                List.ItemType = (list as list) =>
                    let
                        ItemTypes = List.Transform(
                            list,
                            each
                                if Value.Type(Value.FromText(_, culture)) = type number then
                                    if Text.Contains(Text.From(_, culture), "%") then
                                        Percentage.Type
                                    else if Text.Length(Text.Remove(Text.From(_, culture), {"0".."9"} & Text.ToList("., -+eE()/'"))) > 0 then
                                        Currency.Type
                                    else if Int64.From(_, culture) = Value.FromText(_, culture) then
                                        Int64.Type
                                    else
                                        type number
                                else
                                    Value.Type(Value.FromText(_, culture))
                        ),
                        ListItemType = Type.NonNullable(Type.Union(ItemTypes))
                    in
                        ListItemType,
                TypedTable = Table.TransformColumnTypes(table, TransformList, culture)
            in
                TypedTable
    in
        Origine
