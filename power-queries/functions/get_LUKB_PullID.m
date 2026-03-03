// Query   : get_LUKB_PullID
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-03-03T17:39:11+00:00

let
        PullID = (StartURL as text) =>
            let
                html_content = Text.FromBinary(Web.Contents(StartURL)),
                // Find pullID regardless of quotes or spacing
                pull_id_raw = Text.BetweenDelimiters(html_content, "pullID", "'", {0, RelativePosition.FromStart}, 0),
                // Fallback for double quotes
                pull_id_fallback = if Text.Length(pull_id_raw) < 10 then Text.BetweenDelimiters(html_content, "pullID", """", {0, RelativePosition.FromStart}, 0) else pull_id_raw,
                // Clean up separators
                pull_id_clean = Text.Trim(Text.Remove(pull_id_fallback, {":", " ", "=", "'", """"})),
                pull_id = Text.Select(pull_id_clean, {"a".."z", "A".."Z", "0".."9", "_", "-"})
            in
                pull_id
    in
        PullID;

[
    Description = "Gets Data in asingle update_pull request based on a list ofsecuritiesIOs and Data Fields (Data Type, Dimension, Nmmber and format)"
]
