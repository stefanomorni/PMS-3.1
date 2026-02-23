// Query   : lukb_get_json_data
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00

// Power Query function: get_json_data
// Purpose: Fetch JSON data from a web API with retry mechanism and incremental delay
// 
// Parameters:
// - listing_ids: list of listing IDs to fetch data for
// - fields_codes: list of field codes to include in the response
// - max_retries: maximum number of retry attempts
// - in_current_try: (optional) current retry attempt number
// - in_pull_id: (optional) pull ID for the request
// - in_customer_id: (optional) customer ID
// - in_timeout_seconds: (optional) timeout in seconds for the request
// - in_is_retry: (optional) boolean flag indicating if this is a retry attempt

let
    get_json_data = (
        listing_ids as list, 
        fields_codes as list,
        max_retries as number, 
        optional in_current_try as nullable number,
        optional in_pull_id as nullable text, 
        optional in_customer_id as nullable text, 
        optional in_timeout_seconds as nullable number, 
        optional in_is_retry as logical
    ) =>
    let
        // 1. Constants definitions
        updatePull_url = "https://boersenundmaerkte.lukb.ch/lukb/listing/updatePull",
        current_try = if in_current_try = null then 0 else in_current_try,
        is_retry = if in_is_retry = null then true else in_is_retry,

        // 2. Input arguments manipulation
        list_id = if in_customer_id is null
                    then "BFE028A242BD0691902C313CEFA247C9"
                    else in_customer_id,
        
        pull_id = if in_pull_id is null
                    then lukb_get_pull_id(list_id)
                    else in_pull_id,

        referer = "https://boersenundmaerkte.lukb.ch/lukb/lists/CUSTOMER?q=" & list_id & "&t=list",

        timeout = if in_timeout_seconds = null 
                    then #duration(0,0,0,5) // set to 5 seconds; default is 100 seconds
                    else #duration(0,0,0,in_timeout_seconds),

        // 3. Calculate delay for retry
        base_delay = 1,  // Base delay in seconds
        exponential_factor = 1,  // Factor for exponential backoff
        jitter = Number.RandomBetween(0, 500) / 1000,  // Random jitter between 0 and 1 second
        retry_delay = (Number.Power(exponential_factor, current_try) * base_delay) + jitter,

        // 4. Prepare the Web Post request
        headers = [     
            #"User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:128.0) Gecko/20100101 Firefox/128.0",
            #"Accept" = "application/json, text/javascript, */*; q=0.01",
            #"Accept-Language" = "it-CH,en;q=0.5",
            #"Referer" = referer,
            #"X-Requested-With" = "XMLHttpRequest",
            #"Origin" = "https://boersenundmaerkte.lukb.ch",
            #"DNT" = "1",
            #"Sec-GPC" = "1",
            #"Connection" = "keep-alive",
            #"Sec-Fetch-Dest" = "empty",
            #"Sec-Fetch-Mode" = "cors",
            #"Sec-Fetch-Site" = "same-origin"
        ],

        // Prepare the actual web form content data
        post_data = [
            pullID = pull_id, 
            listingIds = listing_ids,
            fields = fields_codes,
            init="true",
            action= "SUBSCRIBE"
        ],  

        // Transform the post_data record to a string
        post_data_string = Uri.BuildQueryString(post_data), 

        // Prepare the "option" record for the pq post requests
        options = [
            Headers = headers,
            IsRetry = is_retry,
            Timeout = timeout,
            Content = Text.ToBinary(post_data_string)
        ],

        // 5. Implement delay before making the request (only for retries)
        delayed_request = if current_try > 0 then Function.InvokeAfter(() => Web.Contents(updatePull_url, options), #duration(0, 0, 0, retry_delay)) else Web.Contents(updatePull_url, options),

        // 6. Make the post request and check status and length
        json_data = Json.Document(try delayed_request otherwise error "Post request failed"),

        // 7. Retry logic
        current_retry = current_try + 1,
        valid_json = 
            if current_retry <= max_retries then
                if Value.Is(json_data, type record) then
                    json_data
                else
                    @get_json_data(listing_ids, fields_codes, max_retries, current_retry, in_pull_id, in_customer_id, in_timeout_seconds, true)
            else 
                error "No valid JSON data was found. Try increasing maximum retries"
    in 
        valid_json
in 
    get_json_data
