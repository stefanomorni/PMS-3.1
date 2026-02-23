// Query   : get_LUKB_list_data
// Category: functions
// Source  : PMS 3.1.xlsm_PowerQuery.m
// Split   : 2026-02-23T16:13:37+00:00
// Note    : [ Description = "Gets Data in asingle update_pull request based on a list ofsecuritiesIOs and Data Fields (Data Type, Dimension, Nmmber and format)" ]

let
    Data = (securityIDs as list, Fields as nullable list, StartURL as nullable text) => 
    let
        UPDATE_PULL_URL = "https://boersenundmaerkte.lukb.ch/lukb/listing/updatePull",

        PULL_HEADERS = [
            #"host" = "boersenundmaerkte.lukb.ch",
            #"Accept" = "application/json, text/javascript, */*; q=0.01",
            #"Referer" = "https://boersenundmaerkte.lukb.ch/lukb/lists/CUSTOMER?q=7C0A8CCE0EB3637A173FE2933873FC05&t=list",
            #"Content-Type" = "application/x-www-form-urlencoded; charset=UTF-8",
            #"Origin" = "https://boersenundmaerkte.lukb.ch"
        ],
        
        BASIC_FIELDS = {
            "LVAL_NORM:value:2:r",
            "LVAL_NORM:datetime:2:c",
            "NC2_PR:value:2:r",
            "NC2_NORM:value:2:r",
            "YTD_PR_NORM:value:2:r"                        
        },
        
        pull_securityIDs = Json.FromValue(List.Transform(securityIDs, each [id = Text.From(_), selected = "true"])),

        pull_fields = if Fields <> null 
                      then Json.FromValue(List.Transform(Fields, each [id = Text.From(_), selected = "true"]))
                      else Json.FromValue(List.Transform(BASIC_FIELDS, each [id = _, selected = "true"])),
        
        URL = if StartURL <> null 
              then Text.Trim(StartURL) 
              else "https://boersenundmaerkte.lukb.ch/lukb/details/1200526,847,1#Tab0",

        ClientPullID = get_LUKB_PullID(URL),       

        pull_data = [
            pullID = ClientPullID, 
            securityIDs = pull_securityIDs,
            fields = pull_fields
        ],
        
        queryString = Uri.BuildQueryString(pull_data),
        
        response = Json.Document(Web.Contents(UPDATE_PULL_URL, [
            Headers = PULL_HEADERS,
            Content = Text.ToBinary(queryString)
        ])),
        data = response[data]
    in
        data
in
    Data
