// Query   : get_LUKB_PullID
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

let 
    PullID = (StartURL as text) => 
        let
            #"Website HTML Lines" = Table.FromColumns({Lines.FromBinary(Web.Contents(StartURL))}),
            #"Filtered Rows" = Table.SelectRows(#"Website HTML Lines", each Text.Contains([Column1], "pullID")),
            #"Split Column by Delimiter" = Table.SplitColumn(#"Filtered Rows", "Column1", Splitter.SplitTextByDelimiter(":", QuoteStyle.Csv), {"Column1.1", "Column1.2"}),
            #"Changed Type" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Column1.1", type text}, {"Column1.2", type text}}),
            #"Column1 1" = #"Changed Type"{0}[Column1.2]
        in
            #"Column1 1"
in
    PullID
