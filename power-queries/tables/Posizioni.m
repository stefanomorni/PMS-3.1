// Query   : Posizioni
// Category: tables
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

let
    smo_lukb_ids_table = smo_lukb_ids_tbl,
    smo_lukb_ids_data_table = lukb_get_smo_ids_data(smo_lukb_ids_tbl[LUKB_id], "fields_types",50,null, null, 10),
    data_fields_list= Record.ToList(Table.DemoteHeaders(smo_lukb_ids_data_table){0}),
    joined_table = Table.NestedJoin(smo_lukb_ids_table, {"LUKB_id"}, smo_lukb_ids_data_table, {"listing_id"}, "smo_lukb_ids_data", JoinKind.LeftOuter),
    expanded_table = Table.ExpandTableColumn(joined_table, "smo_lukb_ids_data", data_fields_list),
    #"Removed Columns" = Table.RemoveColumns(expanded_table,{"listing_id"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns","-",null,Replacer.ReplaceValue,{"BID:volume", "ASK:volume"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Replaced Value",{{"NSIN_CH", type text}, {"BID:volume", Int64.Type}, {"ASK:volume", Int64.Type}, {"DIVIDEND", type number}, {"YIELD", type number},{"YLDEQ", type number}, {"PRICINGAMT", type number}, {"Bid-Ask Spread %", type number}, {"DIVIDEND:datetime", type date}, {"1STPAYMENTYEAR", type datetime}})
in
    #"Changed Type"
