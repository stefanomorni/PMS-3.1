// Query   : lukb_get_pull_id
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-03-03T17:39:11+00:00

let
    PullID = (optional customer_id as nullable text) =>
        let
            base_url = "https://boersenundmaerkte.lukb.ch/lukb/",
            relative_path = "lists/CUSTOMER",
            list_id = if customer_id = null then "BFE028A242BD0691902C313CEFA247C9" else customer_id,
            query = [
                q = list_id,
                t = "list"
            ],
            is_retry = false,
            // ensures that any existing response in the cache is ignored
            timeout = #duration(0, 0, 0, 10),
            // set to 10 seconds; defualt is 100 seconds
            timestamp = DateTime.LocalNow(),
            // Content= [ ] would be used to create a post request
            headers = [
                #"Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
                #"Accept-Language" = "it-CH,en;q=0.5",
                #"Referer" = "https://boersenundmaerkte.lukb.ch/lukb/",
                #"DNT" = "1",
                #"Sec-GPC" = "1",
                #"Connection" = "keep-alive",
                #"Upgrade-Insecure-Requests" = "1",
                #"Sec-Fetch-Dest" = "document",
                #"Sec-Fetch-Mode" = "navigate",
                #"Sec-Fetch-Site" = "same-origin",
                #"Sec-Fetch-User" = "?1",
                #"Priority" = "u=1"
                //#"Timestamp" = DateTime.ToText(timestamp,[Format="yyyy-MM-dd HH:mm:ss.fffffff"])
            ],
            // Return the binary content othe web page; add
            post_response = () =>
                let
                    response = Web.Contents(
                        base_url,
                        [
                            RelativePath = relative_path,
                            Headers = headers,
                            Query = query,
                            IsRetry = is_retry,
                            Timeout = timeout
                        ]
                    )
                in
                    response,
            html_content = Text.FromBinary(post_response()),
            // Find pullID regardless of quotes or spacing
            pull_id_raw = Text.BetweenDelimiters(html_content, "pullID", "'", {0, RelativePosition.FromStart}, 0),
            // Fallback for double quotes if single quotes not found
            pull_id_fallback =
                if Text.Length(pull_id_raw) < 10 then
                    Text.BetweenDelimiters(html_content, "pullID", """", {0, RelativePosition.FromStart}, 0)
                else
                    pull_id_raw,
            // Clean up separators
            pull_id_clean = Text.Trim(Text.Remove(pull_id_fallback, {":", " ", "=", "'", """"})),
            pull_id = Text.Select(pull_id_clean, {"a".."z", "A".."Z", "0".."9", "_", "-"})
        in
            pull_id
in
    PullID
