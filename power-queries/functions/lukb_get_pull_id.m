// Query   : lukb_get_pull_id
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

let 
    PullID=(optional customer_id as nullable text) => 
    let
        base_url = "https://boersenundmaerkte.lukb.ch/lukb/",
        relative_path = "lists/CUSTOMER",
        list_id = if customer_id = null 
                    then "BFE028A242BD0691902C313CEFA247C9" 
                    else customer_id,
        query = [
                q = list_id,
                t = "list"
            ],  
        is_retry = false, // ensures that any existing response in the cache is ignored
        timeout = #duration(0,0,0,2), // set to 2 seconds; defualt is 100 seconds
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
                response = Web.Contents(base_url, 
                    [
                    RelativePath = relative_path,
                    Headers=headers,
                    Query = query,
                    IsRetry = is_retry,
                    Timeout = timeout
                    ])
            in 
                response,

        html_lines = Table.FromColumns({Lines.FromBinary(post_response())}),
        filtered_rows = Table.SelectRows(html_lines, each Text.Contains([Column1], "pullID")),
        split_by_column = Table.SplitColumn(filtered_rows, "Column1", Splitter.SplitTextByDelimiter(":", QuoteStyle.Csv), {"Column1.1", "Column1.2"}),
        pull_id_string = split_by_column{0}[Column1.2],
        pull_id = Text.Trim(pull_id_string)
    in
        pull_id
in
    PullID
