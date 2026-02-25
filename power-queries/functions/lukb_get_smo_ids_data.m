// Query   : lukb_get_smo_ids_data
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

(
    listing_ids_list as list,
    fields_codes_named_range as text,
    max_retries as number,
    optional in_pull_id as nullable text,
    optional customer_id as nullable text,
    optional timeout_seconds as nullable number
) =>
    let
        // The coonnection does sometime block. A call to get_lukb_id will unblock it.
        // reset_connection_id = lukb_get_pull_id(),
        // Collect original input lists from active workbook named ranges
        original_listing_ids = listing_ids_list,
        original_fields_types = fn_range_to_list(fields_codes_named_range),
        // Validate and correct listing_ids by removing duplicates, null values, and invalid formats
        validated_listing_ids = List.Select(List.Distinct(original_listing_ids), each _ <> null),
        corrected_listing_ids = List.Select(
            List.Transform(validated_listing_ids, lukb_validate_listing_id), each _ <> null
        ),
        // Validate and correct fields_types by removing duplicates and null values
        corrected_fields_types = List.Distinct(List.Select(original_fields_types, each _ <> null)),
        // Pass input to functions and retry json collection with 1.5 timeout if it fails
        data = lukb_get_json_data(
            corrected_listing_ids, corrected_fields_types, max_retries, null, in_pull_id, customer_id,
            timeout_seconds, null
        ),
        data_tbl = lukb_json_to_table(data, corrected_listing_ids, corrected_fields_types)
    in
        data_tbl
