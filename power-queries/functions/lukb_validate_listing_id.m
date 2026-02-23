// Query   : lukb_validate_listing_id
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

// Function to validate listing_ids to check if a listing ID is valid (contains three non-empty parts separated by commas
    let validate_listing_id = (id) =>
        let
            parts = Text.Split(id, ","),
            isValid = List.Count(parts) = 3 and List.AllTrue(List.Transform(parts, each Text.TrimStart(_) <> ""))
        in
            if isValid then id else null
    in
        validate_listing_id
