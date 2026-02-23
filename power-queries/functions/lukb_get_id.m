// Query   : lukb_get_id
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

let
lukb_id = (smo_id as text, optional primary_exchange as nullable logical) =>

    let
        // Function parameters example
        fields_types = {
            "M_CUR:value",
            "M_NAME:value",
            "MARKET:value",
            "LVAL:datetime",
            "PMTURNOVER:value"
        },

        // Parse smo_id coomponents for translation into lukb identifiers
        smo_id_elements = Text.Split(smo_id, ":"),
        SearchTerm = try smo_id_elements{0} otherwise null,
        currency_code = try smo_id_elements{1} otherwise null,
        exchange_code = try smo_id_elements{2} otherwise null,

        // Check if is primary exchange based on optional argumento or exchenge = "P" or "p"
        is_primary_exchange = if primary_exchange = true or Text.Lower(exchange_code) = "p"
                                then true  
                                else false,

        // get the corresponding lukb ids for currency and exchange if provided

        CurrencyCode = if currency_code = null 
                        then currency_code 
                        else try Table.SelectRows(lukb_currencies_tbl, each ([cur_code] = currency_code)){0}[cur_id] 
                            otherwise "Error: the currency code was not found",


        AssetClass = null,
        ExchangeCode = if exchange_code = null or Text.Length(exchange_code)= 1
                        then null 
                        else try Table.SelectRows(lukb_exchanges_tbl, each ([exch_code] = exchange_code)){0}[exch_id] 
                            otherwise "Error: the exchange code was not found",




        Search_str= Text.Trim(SearchTerm),
        Currency_str = Text.Trim(CurrencyCode),
        fields_str = Text.Combine(fields_types,","),
        Cur_Str = if CurrencyCode <> null 
                    then Text.Trim(CurrencyCode) 
                    else "",
        AC_str =  if AssetClass <> null and AssetClass <> "" 
                    then Text.Trim(AssetClass) 
                    else "EQU,IND,BON,ETF,FON,TFO,CUR,COM,INT,OPT,DER,FUT,TER,CTR,DIV",
        Exch_str =  if ExchangeCode <> null 
                    then Text.Trim(ExchangeCode) 
                    else "",
        url = "https://boersenundmaerkte.lukb.ch/lukb/api/search",
        headers = [#"Content-Type" = "application/x-www-form-urlencoded;charset=UTF-8"],
        postData = [searchTerm = Search_str, 
                    searchType = "",
                    currencies = Cur_Str,
                    markets =  Exch_str,
                    inputType = "",
                    size = "100",
                    useWildcards = "true",
                    flavor= AC_str,
                    fields = fields_str
                    ],
        response = Json.Document(Web.Contents(url,
            [
                Headers = headers,
                Content = Text.ToBinary(Uri.BuildQueryString(postData))
            ])),
        data = response[data],
        solidListings = data[solidListings],
        solidListingTable = Table.FromList(solidListings, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
        expand_columns = Table.ExpandRecordColumn(solidListingTable, "Column1", Record.FieldNames(solidListingTable[Column1]{0})),
        data_fields = fn_expand_all_records(expand_columns, null, null, true, ":"),
        Filter_exchange = if Exch_str <> "" 
                            then Table.SelectRows(data_fields, each ([#"M_CUR:value"] = Currency_str and [#"MARKET:value"] = Exch_str))
                            else Table.SelectRows(data_fields, each ([#"M_CUR:value"] = Currency_str)),
        //Filter_name = if AC_str = "CUR"  
        //                    then Table.SelectRows(Filter_exchange, each ([#"M_NAME:value"] = Search_str))
        //                    else Filter_exchange,
        remove_value_from_column_names = List.Transform(Table.ColumnNames(data_fields), each {_, Text.Replace(_,":value","")}),
        renamed_value_columns_tbl = Table.RenameColumns(data_fields, remove_value_from_column_names),
        columns_to_remove_list = List.Select(Table.ColumnNames(renamed_value_columns_tbl), each Text.Contains(_, "i_status")),
        remove_status_col_tbl = Table.RemoveColumns(renamed_value_columns_tbl, columns_to_remove_list),
        replace_invalid_values = Table.ReplaceValue(remove_status_col_tbl,"-",null,Replacer.ReplaceValue,{"PMTURNOVER", "LVAL:datetime"}),
        removed_tsd_separator = Table.ReplaceValue(replace_invalid_values,"'","",Replacer.ReplaceText,{"PMTURNOVER"}),
        turnover_to_number = Table.TransformColumnTypes(removed_tsd_separator,{{"PMTURNOVER", type number}}),
        final_table =turnover_to_number,
        datetime_to_datetime = if Exch_str = "" 
                                then Table.TransformColumnTypes(final_table,{{"LVAL:datetime", type datetime}})
                                else final_table,
        sort_by_datetime = if Exch_str = "" 
                            then Table.Sort(datetime_to_datetime,{{"LVAL:datetime", Order.Descending}})
                            else final_table,
        most_current = if Exch_str = "" 
                            then Table.FirstN(sort_by_datetime,1)
                            else final_table,
        sort_by_turnover =if Exch_str = "" 
                            then Table.Sort(datetime_to_datetime,{{"PMTURNOVER", Order.Descending}})
                            else final_table,
        highest_turnover = if Exch_str = "" 
                            then Table.FirstN(sort_by_turnover,1)
                            else final_table,
        result = if is_primary_exchange = true 
                    then highest_turnover
                    else most_current,
        first_result =result[id]{0}
    in
        first_result
in  
    lukb_id
