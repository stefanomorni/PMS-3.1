// Query   : smo_lukb_ids_tbl
// Category: lookups
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-24T17:03:57+00:00

let
        Source = Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "Valori"]}[Content], {"smo_id", "LUKB_id"}),
        ids_to_text = Table.TransformColumnTypes(Source, {{"smo_id", type text}, {"LUKB_id", type text}}),
        zeros_to_null = Table.ReplaceValue(ids_to_text, "0", null, Replacer.ReplaceValue, {"LUKB_id"}),
        errors_to_null = Table.ReplaceErrorValues(zeros_to_null, {{"LUKB_id", null}}),
        import_new_ids = Table.NestedJoin(
            errors_to_null, {"smo_id"}, new_ids_tbl, {"smo_id"}, "new_ids", JoinKind.LeftOuter
        ),
        expand_new_ids = Table.ExpandTableColumn(import_new_ids, "new_ids", {"lukb_id"}, {"lukb_id"}),
        smo_lukb_ids = Table.CombineColumns(
            expand_new_ids, {"LUKB_id", "lukb_id"}, Combiner.CombineTextByDelimiter("", QuoteStyle.None), "LUKB_id"
        )
    in
        smo_lukb_ids
