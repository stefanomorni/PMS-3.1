// Query   : new_ids_tbl
// Category: lookups
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Source = Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "Valori"]}[Content], {"smo_id", "LUKB_id"}),
        ids_to_text = Table.TransformColumnTypes(Source, {{"smo_id", type text}, {"LUKB_id", type text}}),
        zeros_to_null = Table.ReplaceValue(ids_to_text, "0", null, Replacer.ReplaceValue, {"LUKB_id"}),
        errors_to_null = Table.ReplaceErrorValues(zeros_to_null, {{"LUKB_id", null}}),
        missing_ids = Table.SelectRows(errors_to_null, each ([LUKB_id] = null)),
        retrieved_lukb_ids = Table.AddColumn(missing_ids, "lukb_id", each lukb_get_id([smo_id], true))
    in
        retrieved_lukb_ids
