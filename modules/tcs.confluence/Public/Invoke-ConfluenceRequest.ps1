function Invoke-ConfluenceRequest {
    <#
    .SYNOPSIS
        Sends a request to the Confluence REST API using the current Confluence context.

    .DESCRIPTION
        Invoke-ConfluenceRequest is the transport used by every tcs.confluence REST command. It builds
        the endpoint from -Resource (or an explicit -URIPath), appends query parameters and an optional
        CQL title search, sends the request with the credential stored by Set-ConfluenceContext and
        follows "next" pagination links for GET requests up to -MaxQueryPages pages.

        Behaviour that is applied automatically:
        - A wildcard (* or ?) -Search against v2 pages switches to the v1 content API, which supports CQL.
        - A spaceKey query value for v2 pages is resolved to a spaceId (numeric keys are used as-is).
        - Legacy short paths (/pages, /content, /spaces) are mapped to their full API paths.
        - Pagination links pointing to a different host are not followed, so the credential is only
          ever sent to the Confluence site in the context.

        HTTP error responses are reported as errors that include the status and the Confluence error
        titles; use -ErrorAction Stop to turn them into terminating errors.

        Returns an object with a Results array (all collected results) and a MultiPage flag.

    .PARAMETER Method
        The HTTP method: GET, POST, PUT or DELETE.

    .PARAMETER URIPath
        An explicit API path such as /wiki/api/v2/pages. Used as-is unless -Resource is given without
        -URIPath.

    .PARAMETER Resource
        A resource shortcut used to build the path: pages, content or spaces.

    .PARAMETER ApiVersion
        The API version used with -Resource: 2 (default, /wiki/api/v2) or 1 (/wiki/rest/api).

    .PARAMETER Id
        A resource ID appended to the path built from -Resource, for example a page ID.

    .PARAMETER RawPath
        Do not build the path from -Resource; use -URIPath exactly as given.

    .PARAMETER Body
        The JSON request body for POST and PUT requests.

    .PARAMETER Query
        A hashtable of query-string parameters, for example @{ limit = 25; spaceKey = 'DOCS' }.
        Values are URL-encoded.

    .PARAMETER MaxQueryPages
        The maximum number of result pages to retrieve for GET requests. Default 3.

    .PARAMETER Search
        A page title to search for. Wildcards (* ?) produce a CQL "title ~" search, otherwise an exact
        "title =" search.

    .EXAMPLE
        Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = 'DOCS'; limit = 50 }

        Returns up to three pages of results for the pages in the DOCS space.

    .EXAMPLE
        Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'Release*'

        Searches for pages whose title starts with "Release" using the v1 CQL search.

    .EXAMPLE
        Invoke-ConfluenceRequest -Method DELETE -Resource pages -Id 123456 -ErrorAction Stop

        Deletes page 123456 and throws if Confluence reports an error.

    .OUTPUTS
        System.Management.Automation.PSCustomObject with Results and MultiPage properties.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "The API method to use for the request. Valid values: 'GET', 'POST', 'PUT', 'DELETE'.")]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE')]
        [string]$Method,

        [Parameter(HelpMessage = 'Explicit URI path (takes precedence unless -Resource used without -URIPath).')]
        [string]$URIPath,

        [Parameter(HelpMessage = 'High-level resource shortcut.')]
        [ValidateSet('pages', 'content', 'spaces')]
        [string]$Resource,

        [Parameter(HelpMessage = 'API version for -Resource. 1=legacy, 2=new. Default 2.')]
        [ValidateSet(1, 2)]
        [int]$ApiVersion = 2,

        [Parameter(HelpMessage = 'Optional resource Id when using -Resource.')]
        [string]$Id,

        [Parameter(HelpMessage = 'Bypass smart path building even if -Resource supplied.')]
        [switch]$RawPath,

        [Parameter(HelpMessage = 'The body of the request.')]
        [string]$Body,

        [Parameter(HelpMessage = 'Query-string parameters.')]
        [hashtable]$Query,

        [Parameter(HelpMessage = 'Maximum number of result pages to retrieve.')]
        [int16]$MaxQueryPages = 3,

        [Parameter(HelpMessage = 'The search string to use for the request. Using CQL syntax.')]
        [string]$Search
    )

    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        # Helper for URL encoding (portable; avoids System.Web dependency)
        function ConvertTo-UrlEncodedValue([string]$Value) {
            if ($null -eq $Value) { return '' }
            return [System.Net.WebUtility]::UrlEncode($Value)
        }

        # Work on a copy so the caller's hashtable is not changed (spaceKey -> spaceId rewrite)
        if ($Query) { $Query = $Query.Clone() } else { $Query = @{} }

        # Build URIPath from -Resource unless RawPath or explicit URIPath already provided ---
        if (-not $RawPath.IsPresent -and -not $URIPath -and $Resource) {
            switch ($ApiVersion) {
                2 {
                    switch ($Resource) {
                        'pages' { $URIPath = "/wiki/api/v2/pages" }
                        'spaces' { $URIPath = "/wiki/api/v2/spaces" }
                        'content' { Write-Verbose "Resource 'content' with v2 not broadly supported; defaulting to v1 /content."; $ApiVersion = 1; $URIPath = "/wiki/rest/api/content" }
                    }
                }
                1 {
                    switch ($Resource) {
                        'pages' { $URIPath = "/wiki/rest/api/content"; Write-Verbose "Mapping resource 'pages' to v1 /content." }
                        'content' { $URIPath = "/wiki/rest/api/content" }
                        'spaces' { $URIPath = "/wiki/rest/api/space" }
                    }
                }
            }
            if ($Id) {
                $URIPath = ($URIPath.TrimEnd('/')) + "/$Id"
            }
            Write-Verbose "Constructed URIPath from Resource/ApiVersion: $URIPath"
        } elseif (-not $URIPath) {
            Write-Error "Provide either -URIPath or -Resource."
            return
        }

        $ConfluenceContext = $script:ConfluenceContext
        if ($null -eq $ConfluenceContext -or $null -eq $script:ConfluenceCredential) {
            Write-Error 'No Confluence context is set. Run Set-ConfluenceContext first.'
            return
        }
        Write-Verbose ("Raw ConnectionURI object(s): {0}" -f ($ConfluenceContext.ConnectionURI -join ', '))

        # --- Enhanced base normalization (strip api suffix if present) ---
        $rawBase = ($ConfluenceContext.ConnectionURI | ForEach-Object { "$($_)" }) -join ''
        $rawBase = $rawBase.Trim()
        if (-not $rawBase) { Write-Error "Empty ConnectionURI (after trimming)."; return }
        if ($rawBase.StartsWith('"') -and $rawBase.EndsWith('"')) { $rawBase = $rawBase.Trim('"') }
        if ($rawBase.StartsWith("'") -and $rawBase.EndsWith("'")) { $rawBase = $rawBase.Trim("'") }
        $rawBase = ($rawBase -replace '\s+', '')
        $rawBase = ($rawBase -replace '^(https?://)+', '$1')

        if ($rawBase -notmatch '^[a-zA-Z][a-zA-Z0-9+\-.]*://') {
            $rawBase = "https://$rawBase"
            Write-Verbose "Added https:// scheme to base URI candidate."
        }

        # Strip trailing known API suffixes up-front to avoid compounded duplication later
        $rawBaseNoApi = $rawBase -replace '(?i)/(?:wiki/)?(?:api/v2|rest/api)/*$', ''
        if ($rawBaseNoApi -ne $rawBase) {
            Write-Verbose "Trimmed API suffix from base URI ('$rawBase' -> '$rawBaseNoApi')."
            $rawBase = $rawBaseNoApi
        }

        $baseUriObj = $null
        if (-not [System.Uri]::TryCreate($rawBase, [System.UriKind]::Absolute, [ref]$baseUriObj)) {
            Write-Error "Failed to parse ConnectionURI after normalization attempt. Value: '$rawBase'"
            return
        }
        if ($baseUriObj.Query) { Write-Verbose "Stripping query part from base URI." }
        if ($baseUriObj.Fragment) { Write-Verbose "Stripping fragment part from base URI." }

        $BaseUriString = "{0}://{1}{2}" -f $baseUriObj.Scheme, $baseUriObj.Authority, ($baseUriObj.AbsolutePath.TrimEnd('/'))
        if (-not $BaseUriString) { Write-Error "Failed to reconstruct normalized base URI."; return }
        Write-Verbose "Normalized base URI (pre path merge): $BaseUriString"

        try {
            $OriginalBaseUriString = $BaseUriString
            $BaseUriString = $BaseUriString.TrimEnd('/')
            Write-Verbose "BaseUri normalized from '$OriginalBaseUriString' to '$BaseUriString'"
            $URIPathNormalized = $URIPath
            Write-Verbose "Original URIPath: '$URIPath'"
            if (-not $URIPathNormalized.StartsWith('/')) { $URIPathNormalized = "/$URIPathNormalized" }

            if ($BaseUriString -match '/wiki$' -and $URIPathNormalized -like '/wiki/*') {
                Write-Verbose "Detected duplicate /wiki segment. Removing leading /wiki from path."
                $URIPathNormalized = $URIPathNormalized.Substring(5)
            }

            # Legacy simple path normalization (auto-upgrade older function calls)
            switch ($URIPathNormalized.ToLower()) {
                '/pages' {
                    Write-Verbose "Legacy path '/Pages' mapped to '/wiki/api/v2/pages'."
                    $URIPathNormalized = '/wiki/api/v2/pages'
                }
                '/content' {
                    Write-Verbose "Legacy path '/Content' mapped to '/wiki/rest/api/content'."
                    $URIPathNormalized = '/wiki/rest/api/content'
                }
                '/spaces' {
                    Write-Verbose "Legacy path '/Spaces' mapped to '/wiki/api/v2/spaces'."
                    $URIPathNormalized = '/wiki/api/v2/spaces'
                }
            }

            # --- Automatic wildcard / API version handling ---
            $UsingV2Pages = $false
            if ($URIPathNormalized -match '^/wiki/api/v2/pages' -or $URIPathNormalized -match '^/api/v2/pages') {
                if ($URIPathNormalized -match '^/api/v2/') { $URIPathNormalized = "/wiki$URIPathNormalized" }
                $UsingV2Pages = $true
            }
            $WildcardSearch = ([string]::IsNullOrWhiteSpace($Search) -eq $false -and $Search -match '[\*\?]')
            if ($UsingV2Pages -and $WildcardSearch) {
                Write-Verbose "Wildcard search detected for v2 pages path. Switching endpoint to v1 content API for CQL compatibility."
                if ($URIPathNormalized -match '^/wiki/api/v2/pages/(\d+)$') {
                    $pageId = $Matches[1]
                    $URIPathNormalized = "/wiki/rest/api/content/$pageId"
                } else {
                    $URIPathNormalized = "/wiki/rest/api/content"
                }
                $UsingV2Pages = $false
            }

            $EndpointBase = "$BaseUriString$URIPathNormalized"
            $EndpointBeforeCollapse = $EndpointBase
            $EndpointBase = [regex]::Replace($EndpointBase, '/wiki/(api/v2|rest/api)/wiki/\1', '/wiki/$1')
            if ($EndpointBase -ne $EndpointBeforeCollapse) { Write-Verbose "Collapsed duplicated API path segment in endpoint base." }
            Write-Verbose "Normalized URIPath: '$URIPathNormalized'"
            Write-Verbose "Endpoint base (pre query): $EndpointBase"
        }
        catch {
            Write-Error "Invalid URL components. Base: '$($ConfluenceContext.ConnectionURI)' Path: '$URIPath'. Error: $_"
            return
        }

        # --- spaceKey / spaceId handling for v2 pages ---
        if ($UsingV2Pages -and $Query.Count -gt 0) {
            if ($Query.ContainsKey('spaceKey')) {
                $spaceKeyVal = $Query.spaceKey
                if ($spaceKeyVal -match '^\d+$') {
                    Write-Verbose "spaceKey value '$spaceKeyVal' numeric; treating as spaceId."
                    $Query.Remove('spaceKey') | Out-Null
                    $Query.spaceId = $spaceKeyVal
                } elseif ($spaceKeyVal) {
                    Write-Verbose "Resolving spaceKey '$spaceKeyVal' to spaceId (v2)."
                    try {
                        $spaceLookup = Invoke-ConfluenceRequest -Method GET -URIPath "/wiki/api/v2/spaces" -Query @{ keys = $spaceKeyVal } -MaxQueryPages 1 -RawPath -ErrorAction Stop
                        $resolvedSpaceId = ($spaceLookup.Results | Where-Object { $_.key -eq $spaceKeyVal }).id
                        if ($resolvedSpaceId) {
                            Write-Verbose "spaceKey '$spaceKeyVal' resolved to spaceId '$resolvedSpaceId'."
                            $Query.Remove('spaceKey') | Out-Null
                            $Query.spaceId = $resolvedSpaceId
                        } else {
                            Write-Warning "spaceKey '$spaceKeyVal' could not be resolved; leaving spaceKey parameter."
                        }
                    } catch {
                        Write-Warning "Failed to resolve spaceKey '$spaceKeyVal' to spaceId. Error: $_"
                    }
                }
            }
        }

        # --- Build query string / CQL (reworked) ---
        $QueryParts = @()
        if ($Query.Count -gt 0) {
            # get_Keys(): a query parameter named "keys" would otherwise hide the Keys property
            $queryKeys = @($Query.get_Keys())
            Write-Verbose ("Processing query hashtable keys: {0}" -f ($queryKeys -join ', '))
            foreach ($k in $queryKeys) {
                $QueryParts += ("{0}={1}" -f (ConvertTo-UrlEncodedValue $k), (ConvertTo-UrlEncodedValue ([string]$Query[$k])))
            }
        }
        if ([string]::IsNullOrWhiteSpace($Search) -eq $false) {
            $HasWild = $Search -match '[\*\?]'
            $CqlOperator = $(if ($HasWild) { "~" } else { "=" })
            $EscapedSearch = $Search.Replace('"', '\"')
            $Cql = "title $CqlOperator `"$EscapedSearch`""
            Write-Verbose "Generated CQL: $Cql"
            $QueryParts += ("cql={0}" -f (ConvertTo-UrlEncodedValue $Cql))
        }

        # Deterministic final endpoint assembly
        $Endpoint = $EndpointBase
        if (-not $Endpoint) {
            Write-Verbose "EndpointBase empty; reconstructing from BaseUriString + URIPathNormalized."
            $EndpointBase = "$BaseUriString$URIPathNormalized"
            $Endpoint = $EndpointBase
        }

        if ($QueryParts.Count -gt 0) {
            $queryString = ($QueryParts -join '&')
            $Endpoint = "$EndpointBase`?$queryString"
            Write-Verbose "Full request URI: $Endpoint"
        } else {
            Write-Verbose "Full request URI (no query params): $Endpoint"
        }

        # Repair safeguard: if scheme missing (symptom previously seen)
        if ($Endpoint -notmatch '^[a-zA-Z][a-zA-Z0-9+\-.]*://') {
            Write-Warning "Endpoint lost base scheme/header. Repairing using EndpointBase."
            if ($Endpoint -match '=') {
                $Endpoint = "$EndpointBase`?" + $Endpoint.TrimStart('?')
            } else {
                $Endpoint = $EndpointBase
            }
            Write-Verbose "Repaired full request URI: $Endpoint"
        }

        # Final validation
        $finalUriObj = $null
        if (-not [System.Uri]::TryCreate($Endpoint, [System.UriKind]::Absolute, [ref]$finalUriObj)) {
            Write-Error "Final endpoint URI invalid. Endpoint='$Endpoint' BaseUriString='$BaseUriString' URIPathNormalized='$URIPathNormalized' QueryPartsCount=$($QueryParts.Count)"
            return
        }
        Write-Verbose "Validated endpoint URI: $($finalUriObj.AbsoluteUri)"


        # --- Send the request(s) ---
        $OutputObject = [pscustomobject]@{ Results = @(); MultiPage = $false }
        $QueryPageCount = 0
        $CurrentEndpoint = $Endpoint
        Write-Verbose "Starting request loop. MaxQueryPages=$MaxQueryPages Method=$Method"
        do {
            Write-Verbose ('[Page {0}] {1} {2}' -f ($QueryPageCount + 1), $Method.ToUpper(), $CurrentEndpoint)
            $requestParams = @{
                Uri    = $CurrentEndpoint
                Method = $Method.ToUpper()
            }
            if ($Body) { $requestParams.Body = $Body }
            try {
                $Response = Invoke-ConfluenceHttpRequest @requestParams -ErrorAction Stop
            }
            catch {
                Write-Error "Request to $CurrentEndpoint failed. $($_.Exception.Message)"
                return
            }

            if ($Response.StatusCode -lt 200 -or $Response.StatusCode -gt 299) {
                $Errors = $null
                try {
                    $errContent = $Response.Content | ConvertFrom-Json -ErrorAction Stop
                    $Errors = (@($errContent.errors | ForEach-Object { $_.title }) + @($errContent.message) | Where-Object { $_ }) -join '; '
                }
                catch {
                    Write-Verbose 'The error response body is not JSON.'
                }
                Write-Error "Failed request. Status: $($Response.StatusCode) $($Response.StatusDescription) Errors: $Errors URL: $CurrentEndpoint"
                return
            }

            $ResponsePayload = $null
            if (-not [string]::IsNullOrWhiteSpace($Response.Content)) {
                try {
                    $ResponsePayload = $Response.Content | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed to parse JSON response from $CurrentEndpoint. $_"
                    return
                }
            }
            else {
                Write-Verbose "Empty response body (status $($Response.StatusCode))."
            }

            $NextLink = $null
            if ($null -ne $ResponsePayload) {
                if ($ResponsePayload.PSObject.Properties.Name -contains 'results') {
                    $OutputObject.Results += $ResponsePayload.results
                    $OutputObject.MultiPage = $true
                }
                else {
                    $OutputObject.Results += $ResponsePayload
                }

                if ($ResponsePayload.PSObject.Properties.Name -contains '_links') {
                    $linksObj = $ResponsePayload._links
                    if ($linksObj -and ($linksObj.PSObject.Properties.Name -contains 'next')) {
                        $NextLink = $linksObj.next
                    }
                }
                if (-not $NextLink -and $ResponsePayload.PSObject.Properties.Name -contains 'links') {
                    $linksObj = $ResponsePayload.links
                    if ($linksObj -and ($linksObj.PSObject.Properties.Name -contains 'next')) {
                        $NextLink = $linksObj.next
                    }
                }
            }

            $CurrentEndpoint = $null
            if ($NextLink) {
                if ($NextLink -match '^[a-z]+://') {
                    $nextUriObj = $null
                    if ([System.Uri]::TryCreate($NextLink, [System.UriKind]::Absolute, [ref]$nextUriObj) -and $nextUriObj.Scheme -eq 'https' -and $nextUriObj.Authority -eq $baseUriObj.Authority) {
                        $CurrentEndpoint = $NextLink
                    }
                    else {
                        Write-Warning 'Ignoring a pagination link that points outside the Confluence site in the context.'
                    }
                }
                else {
                    if (-not $NextLink.StartsWith('/')) { $NextLink = "/$NextLink" }
                    # v2 links are relative to the site root, v1 links to the /wiki context path
                    if ($NextLink -notlike '/wiki/*' -and $BaseUriString -notmatch '/wiki$') { $NextLink = "/wiki$NextLink" }
                    $CurrentEndpoint = $BaseUriString.TrimEnd('/') + $NextLink
                }
                if ($CurrentEndpoint) { Write-Verbose "Detected next page link: $CurrentEndpoint" }
            }
            else {
                Write-Verbose 'No further pagination link found.'
            }

            $QueryPageCount++
            if ($QueryPageCount -ge $MaxQueryPages) {
                Write-Verbose 'Maximum query page count reached. Exiting loop.'
                break
            }
        } while ($null -ne $CurrentEndpoint -and $Method -eq 'GET')

        Write-Verbose ('Completed request. Pages retrieved: {0}. Total result objects collected: {1}' -f $QueryPageCount, @($OutputObject.Results).Count)
        return $OutputObject
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        throw
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
