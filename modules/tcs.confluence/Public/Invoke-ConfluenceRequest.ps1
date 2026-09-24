function Invoke-ConfluenceRequest {
    <#
    .SYNOPSIS
        Sends a request to the Confluence REST API using the current Confluence context.

    .DESCRIPTION
        Invoke-ConfluenceRequest is the transport used by every tcs.confluence REST command. It builds
        the endpoint from -Resource (or an explicit -URIPath) on the site set by Set-ConfluenceContext,
        appends query parameters and an optional CQL title search, sends the request with the stored
        credential and follows "next" pagination links for GET requests up to -MaxQueryPages pages
        (or all pages with -All).

        Behaviour that is applied automatically:
        - -Resource uses the API version given by -ApiVersion or, when that is not given, the version
          chosen with Set-ConfluenceContext -ApiVersion (v2 by default).
        - A -Search against the page or content collection is sent to the v1 CQL search endpoint
          /wiki/rest/api/content/search as cql=title = "..." (or title ~ "..." with wildcards * ?).
          Page searches also add type = page. CQL values are escaped (backslash, then quote).
        - A spaceKey query value is resolved to the numeric space ID and sent as the documented
          space-id parameter for v2 pages (keys are looked up once per session and cached), or added
          to the CQL query (space = "KEY" or space.id = 123) for searches. A key that cannot be
          resolved is reported as an error and no request is sent.
        - The short paths /pages, /content and /spaces are mapped to their full API paths.
        - A 429 Too Many Requests response is retried, honouring the Retry-After header.
        - Pagination links pointing to a different host are not followed, so the credential is only
          ever sent to the Confluence site in the context.
        - When -MaxQueryPages stops the paging while more results are available, a warning is written.

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
        The API version used with -Resource: 2 (/wiki/api/v2) or 1 (/wiki/rest/api). When not given,
        the version from Set-ConfluenceContext -ApiVersion is used.

    .PARAMETER Id
        A resource ID appended to the path built from -Resource, for example a page ID.

    .PARAMETER RawPath
        Do not build the path from -Resource or map short paths; use -URIPath exactly as given.

    .PARAMETER Body
        The JSON request body for POST and PUT requests.

    .PARAMETER Query
        A hashtable of query-string parameters, for example @{ limit = 25; spaceKey = 'DOCS' }.
        Values are URL-encoded. The caller's hashtable is not changed.

    .PARAMETER MaxQueryPages
        The maximum number of result pages to retrieve for GET requests. Default 3.

    .PARAMETER All
        Follow the pagination links until every result has been retrieved (ignores -MaxQueryPages).

    .PARAMETER Search
        A page title to search for with CQL. Wildcards (* ?) produce a "title ~" search, otherwise an
        exact "title =" search.

    .EXAMPLE
        Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = 'DOCS'; limit = 50 }

        Returns up to three pages of results for the pages in the DOCS space (sent as space-id).

    .EXAMPLE
        Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'Release*' -All

        Searches every page whose title matches "Release*" through /wiki/rest/api/content/search.

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

        [Parameter(HelpMessage = 'API version for -Resource. 1=legacy, 2=new. Default: the context API version.')]
        [ValidateSet(1, 2)]
        [int]$ApiVersion,

        [Parameter(HelpMessage = 'Optional resource Id when using -Resource.')]
        [string]$Id,

        [Parameter(HelpMessage = 'Use -URIPath exactly as given.')]
        [switch]$RawPath,

        [Parameter(HelpMessage = 'The body of the request.')]
        [string]$Body,

        [Parameter(HelpMessage = 'Query-string parameters.')]
        [hashtable]$Query,

        [Parameter(HelpMessage = 'Maximum number of result pages to retrieve.')]
        [int16]$MaxQueryPages = 3,

        [Parameter(HelpMessage = 'Retrieve every result page.')]
        [switch]$All,

        [Parameter(HelpMessage = 'A page title to search for with CQL; wildcards (* ?) supported.')]
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
        $context = $script:ConfluenceContext

        # --- Path ---
        if (-not $RawPath -and -not $URIPath -and $Resource) {
            if (-not $PSBoundParameters.ContainsKey('ApiVersion')) {
                $ApiVersion = 2
                if ($context -and $context.ApiVersion -eq 'v1') { $ApiVersion = 1 }
            }
            $resourcePaths = @{
                2 = @{ pages = '/wiki/api/v2/pages'; spaces = '/wiki/api/v2/spaces'; content = '/wiki/rest/api/content' }
                1 = @{ pages = '/wiki/rest/api/content'; spaces = '/wiki/rest/api/space'; content = '/wiki/rest/api/content' }
            }
            $URIPath = $resourcePaths[$ApiVersion][$Resource]
            if ($Id) { $URIPath = "$URIPath/$Id" }
        }
        elseif (-not $URIPath) {
            Write-Error 'Provide either -URIPath or -Resource.'
            return
        }
        if ($null -eq $context -or $null -eq $script:ConfluenceCredential) {
            Write-Error 'No Confluence context is set. Run Set-ConfluenceContext first.'
            return
        }

        $path = $URIPath.Trim()
        if (-not $path.StartsWith('/')) { $path = "/$path" }
        if (-not $RawPath) {
            $shortPaths = @{ '/pages' = '/wiki/api/v2/pages'; '/content' = '/wiki/rest/api/content'; '/spaces' = '/wiki/api/v2/spaces' }
            if ($shortPaths.ContainsKey($path.ToLowerInvariant())) { $path = $shortPaths[$path.ToLowerInvariant()] }
        }

        # --- Query and CQL ---
        # Work on a copy so the caller's hashtable is not changed
        $queryValues = @{}
        if ($Query) { foreach ($key in @($Query.get_Keys())) { $queryValues[$key] = $Query[$key] } }
        $cqlClauses = New-Object -TypeName System.Collections.Generic.List[string]

        if (-not $RawPath -and -not [string]::IsNullOrWhiteSpace($Search)) {
            if ($path -match '^/wiki/api/v2/pages/?$') {
                $path = '/wiki/rest/api/content/search'
                $cqlClauses.Add('type = page')
            }
            elseif ($path -match '^/wiki/rest/api/content/?$') {
                $path = '/wiki/rest/api/content/search'
            }
        }
        $isCqlEndpoint = $path -match '^/wiki/rest/api/(content/)?search/?$'
        $isV2 = $path -like '/wiki/api/v2/*'

        $spaceValue = $null
        if ($queryValues.ContainsKey('spaceKey') -and ($isCqlEndpoint -or $isV2)) {
            $spaceValue = "$($queryValues['spaceKey'])"
            $queryValues.Remove('spaceKey')
        }
        if ($isV2 -and $queryValues.ContainsKey('spaceId')) {
            $spaceValue = "$($queryValues['spaceId'])"
            $queryValues.Remove('spaceId')
        }
        if (-not [string]::IsNullOrWhiteSpace($spaceValue)) {
            if ($isCqlEndpoint) {
                if ($spaceValue -match '^\d+$') { $cqlClauses.Add("space.id = $spaceValue") }
                else { $cqlClauses.Add('space = ' + (ConvertTo-ConfluenceCqlString -Value $spaceValue)) }
            }
            else {
                try {
                    $queryValues['space-id'] = Resolve-ConfluenceSpaceId -Space $spaceValue -ErrorAction Stop
                }
                catch {
                    Write-Error "Could not resolve space '$spaceValue' to a space ID; no request was sent. $($_.Exception.Message)"
                    return
                }
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($Search)) {
            $operator = if ($Search -match '[\*\?]') { '~' } else { '=' }
            $cqlClauses.Add("title $operator " + (ConvertTo-ConfluenceCqlString -Value $Search))
        }
        if ($cqlClauses.Count -gt 0) {
            if ($queryValues.ContainsKey('cql') -and $queryValues['cql']) { $cqlClauses.Insert(0, "($($queryValues['cql']))") }
            $queryValues['cql'] = $cqlClauses -join ' AND '
            Write-Verbose "CQL: $($queryValues['cql'])"
        }

        # --- Endpoint ---
        $baseUrl = "$($context.ConnectionBaseURL)".TrimEnd('/')
        $endpoint = $baseUrl + $path
        if ($queryValues.Count -gt 0) {
            # get_Keys(): a query parameter named "keys" would otherwise hide the Keys property
            $queryParts = foreach ($key in (@($queryValues.get_Keys()) | Sort-Object)) {
                '{0}={1}' -f [System.Net.WebUtility]::UrlEncode($key), [System.Net.WebUtility]::UrlEncode([string]$queryValues[$key])
            }
            $endpoint = $endpoint + '?' + ($queryParts -join '&')
        }
        $baseUri = $null
        $endpointUri = $null
        if (-not [System.Uri]::TryCreate($baseUrl, [System.UriKind]::Absolute, [ref]$baseUri) -or
            -not [System.Uri]::TryCreate($endpoint, [System.UriKind]::Absolute, [ref]$endpointUri)) {
            Write-Error "The request URI '$endpoint' is not valid."
            return
        }

        # --- Send the request(s) ---
        $results = New-Object -TypeName System.Collections.Generic.List[object]
        $multiPage = $false
        $pageCount = 0
        $currentEndpoint = $endpoint
        do {
            Write-Verbose ('[Page {0}] {1} {2}' -f ($pageCount + 1), $Method, $currentEndpoint)
            $requestParams = @{ Uri = $currentEndpoint; Method = $Method }
            if ($Body) { $requestParams.Body = $Body }
            try {
                $response = Invoke-ConfluenceHttpRequest @requestParams -ErrorAction Stop
            }
            catch {
                Write-Error "Request to $currentEndpoint failed. $($_.Exception.Message)"
                return
            }

            if ($response.StatusCode -lt 200 -or $response.StatusCode -gt 299) {
                $errorText = $null
                try {
                    $errorContent = $response.Content | ConvertFrom-Json -ErrorAction Stop
                    $errorText = (@($errorContent.errors | ForEach-Object { $_.title }) + @($errorContent.message) | Where-Object { $_ }) -join '; '
                }
                catch {
                    Write-Verbose 'The error response body is not JSON.'
                }
                Write-Error "Failed request. Status: $($response.StatusCode) $($response.StatusDescription) Errors: $errorText URL: $currentEndpoint"
                return
            }

            $payload = $null
            if (-not [string]::IsNullOrWhiteSpace($response.Content)) {
                try {
                    $payload = $response.Content | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed to parse JSON response from $currentEndpoint. $_"
                    return
                }
            }

            $nextLink = $null
            if ($null -ne $payload) {
                $propertyNames = @($payload.PSObject.Properties.Name)
                if ($propertyNames -contains 'results') {
                    foreach ($item in @($payload.results)) { if ($null -ne $item) { $results.Add($item) } }
                    $multiPage = $true
                }
                elseif ($payload -is [array]) {
                    foreach ($item in $payload) { $results.Add($item) }
                }
                else {
                    $results.Add($payload)
                }
                foreach ($linkProperty in @('_links', 'links')) {
                    if (-not $nextLink -and $propertyNames -contains $linkProperty -and $payload.$linkProperty -and
                        @($payload.$linkProperty.PSObject.Properties.Name) -contains 'next') {
                        $nextLink = [string]$payload.$linkProperty.next
                    }
                }
            }
            $pageCount++

            $currentEndpoint = $null
            if ($nextLink) {
                if ($nextLink -match '^[a-z][a-z0-9+.-]*://') {
                    $nextUri = $null
                    if ([System.Uri]::TryCreate($nextLink, [System.UriKind]::Absolute, [ref]$nextUri) -and $nextUri.Scheme -eq 'https' -and $nextUri.Authority -eq $baseUri.Authority) {
                        $currentEndpoint = $nextLink
                    }
                    else {
                        Write-Warning 'Ignoring a pagination link that points outside the Confluence site in the context.'
                    }
                }
                else {
                    if (-not $nextLink.StartsWith('/')) { $nextLink = "/$nextLink" }
                    # v2 links are relative to the site root, v1 links to the /wiki context path
                    if ($nextLink -notlike '/wiki/*') { $nextLink = "/wiki$nextLink" }
                    $currentEndpoint = $baseUrl + $nextLink
                }
            }

            if ($Method -ne 'GET' -or $null -eq $currentEndpoint) { break }
            if (-not $All -and $pageCount -ge $MaxQueryPages) {
                Write-Warning "Stopped after $pageCount result page(s) (-MaxQueryPages $MaxQueryPages); more results are available. Use -All or a higher -MaxQueryPages to get them."
                break
            }
        } while ($true)

        Write-Verbose ('Completed request. Pages retrieved: {0}. Results: {1}' -f $pageCount, $results.Count)
        return [pscustomobject]@{
            Results   = $results.ToArray()
            MultiPage = $multiPage
        }
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
