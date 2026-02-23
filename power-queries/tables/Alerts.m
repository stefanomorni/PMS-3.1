// Query   : Alerts
// Category: tables
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00
let
    // VALIDATION TEST 2026-02-23 (v2) — canonical file check
    Origine = Excel.CurrentWorkbook(){[Name = "Entry_Levels"]}[Content],
    #"Modificato tipo" = Table.TransformColumnTypes(
        Origine,
        {
            {"Market", type text},
            {"Instrument", type text},
            {"Type", type text},
            {"Condition", type text},
            {"Value", type number},
            {"Repeat", type logical},
            {"Active", type logical},
            {"Comment", type text},
            {"Current Value", type number},
            {"Distance to Market", type number},
            {"Pct Distance", type number}
        }
    )
in
    #"Modificato tipo"
