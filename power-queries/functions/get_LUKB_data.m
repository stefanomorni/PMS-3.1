// Query   : get_LUKB_data
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Data = (
            SearchTerm as text,
            CurrencyCode as text,
            Fields as text,
            optional AssetClass as nullable text,
            optional ExchangeCode as nullable text
        ) =>
            let
                //Fields = "M_NAME,ISIN,M_CUR,MARKET,LVAL:datetime:value,PRICINGAMT",
                Search_str = Text.Trim(SearchTerm),
                Currency_str = Text.Trim(CurrencyCode),
                // "M_NAME,ISIN,M_CUR,MARKET,LVAL:datetime:value,PRICINGAMT"
                Fields_str = Text.Replace(Fields, " ", ""),
                Cur_Str = if CurrencyCode <> null then Text.Trim(CurrencyCode) else "",
                AC_str =
                    if AssetClass <> null and AssetClass <> "" then
                        Text.Trim(AssetClass)
                    else
                        "EQU,IND,BON,ETF,FON,TFO,CUR,COM,INT,OPT,DER,FUT,TER,CTR,DIV",
                Exch_str = if ExchangeCode <> null then Text.Trim(ExchangeCode) else "",
                url = "https://boersenundmaerkte.lukb.ch/lukb/api/search",
                headers = [#"Content-Type" = "application/x-www-form-urlencoded;charset=UTF-8"],
                postData = [
                    searchTerm = Search_str,
                    searchType = "",
                    currencies = Cur_Str,
                    markets = Exch_str,
                    inputType = "",
                    size = "100",
                    useWildcards = "true",
                    flavor = AC_str,
                    fields = "M_CUR,MARKET,LVAL:datetime," & Fields_str
                ],
                response = Json.Document(
                    Web.Contents(url, [
                        Headers = headers,
                        Content = Text.ToBinary(Uri.BuildQueryString(postData))
                    ])
                ),
                data = response[data],
                solidListings = data[solidListings],
                solidListingTable = Table.FromList(
                    solidListings, Splitter.SplitByNothing(), null, null, ExtraValues.Error
                ),
                expand_columns = Table.ExpandRecordColumn(
                    solidListingTable, "Column1", Record.FieldNames(solidListingTable[Column1]{0})
                ),
                data_fields = fn_expand_all_records(expand_columns, null, null, true, ":"),
                //Filter_exchange = if Exch_str <> ""
                //                    then Table.SelectRows(data_fields, each ([#"M_CUR:value"] = Currency_str and [#"MARKET:value"] = Exch_str))
                //                    else Table.SelectRows(data_fields, each ([#"M_CUR:value"] = Currency_str)),
                //Filter_name = if AC_str = "CUR"
                //                    then Table.SelectRows(Filter_exchange, each ([#"M_NAME:value"] = Search_str))
                //                    else Filter_exchange,
                remove_value_from_column_names = List.Transform(
                    Table.ColumnNames(data_fields), each {_, Text.Replace(_, ":value", "")}
                ),
                renamed_value_columns_tbl = Table.RenameColumns(data_fields, remove_value_from_column_names),
                columns_to_remove_list = List.Select(
                    Table.ColumnNames(renamed_value_columns_tbl), each Text.Contains(_, "i_status")
                ),
                remove_status_col_tbl = Table.RemoveColumns(renamed_value_columns_tbl, columns_to_remove_list),
                // Sort by most current price if non price source (Exchange) passed
                Datetime_To_DateTime =
                    if Exch_str = "" then
                        Table.TransformColumnTypes(remove_status_col_tbl, {{"LVAL:datetime", type datetime}})
                    else
                        remove_status_col_tbl,
                Sort_By_Datetime =
                    if Exch_str = "" then
                        Table.Sort(Datetime_To_DateTime, {{"LVAL:datetime", Order.Descending}})
                    else
                        Datetime_To_DateTime,
                Kept_Most_Current = if Exch_str = "" then Table.FirstN(Sort_By_Datetime, 1) else Sort_By_Datetime
            in
                Kept_Most_Current
    in
        Data
