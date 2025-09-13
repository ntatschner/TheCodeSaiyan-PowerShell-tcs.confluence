function Invoke-ConfluenceRequest {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "The API method to use for the request. Valid values: 'GET', 'POST', 'PUT', 'DELETE'.")]
        [string]$Method,

        [Parameter(HelpMessage = "Explicit URI path (takes precedence unless -Resource used without -RawPath).")]
        [string]$URIPath,

        [Parameter(HelpMessage = "High-level resource shortcut.")]
        [ValidateSet('pages','content','spaces')]
        [string]$Resource,

        [Parameter(HelpMessage = "API version for -Resource. 1=legacy, 2=new. Default 2.")]
        [ValidateSet(1,2)]
        [int]$ApiVersion = 2,

        [Parameter(HelpMessage = "Optional resource Id when using -Resource.")]
        [string]$Id,

        [Parameter(HelpMessage = "Bypass smart path building even if -Resource supplied.")]
        [switch]$RawPath,

        [Parameter(HelpMessage = "The body of the request.")]
        [string]$Body,

        # FIX: was [psobject]; needs hashtable for .ContainsKey/.Remove/.GetEnumerator usage
        [hashtable]$Query,

        [int16]$MaxQueryPages = 3,

        [Parameter(HelpMessage = "The search string to use for the request. Using CQL syntax.")]
        [string]$Search
    )

    begin {
        # Helper for URL encoding (portable; avoids System.Web dependency)
        function _Encode([string]$v) {
            if ($null -eq $v) { return "" }
            return [System.Net.WebUtility]::UrlEncode($v)
        }

        # Ensure Query is a hashtable
        if (-not $Query) { $Query = @{} }

        # --- New: Build URIPath from -Resource unless RawPath or explicit URIPath already provided ---
        if (-not $RawPath.IsPresent -and -not $URIPath -and $Resource) {
            switch ($ApiVersion) {
                2 {
                    switch ($Resource) {
                        'pages'   { $URIPath = "/wiki/api/v2/pages" }
                        'spaces'  { $URIPath = "/wiki/api/v2/spaces" }
                        'content' { Write-Verbose "Resource 'content' with v2 not broadly supported; defaulting to v1 /content."; $ApiVersion = 1; $URIPath = "/wiki/rest/api/content" }
                    }
                }
                1 {
                    switch ($Resource) {
                        'pages'   { $URIPath = "/wiki/rest/api/content"; Write-Verbose "Mapping resource 'pages' to v1 /content." }
                        'content' { $URIPath = "/wiki/rest/api/content" }
                        'spaces'  { $URIPath = "/wiki/rest/api/space" }
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

        $InitialWasV2Pages = $false
        if ($URIPath -match '^/?wiki/api/v2/pages') { $InitialWasV2Pages = $true }

        if ( -Not (Get-Variable -Scope Global -Name "ConfluenceContext" -ErrorAction SilentlyContinue)) {
            Write-Error "ConfluenceContext variable not found. Please run Set-ConfluenceContext to set the token."
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

            if ($URIPathNormalized -match '^/(?:wiki/)?(api/v2|rest/api)(/.*)?$') {
                # Keep as is
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

            # REPLACED: old $Endpoint construction block
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
            Write-Verbose ("Processing query hashtable keys: {0}" -f ($Query.Keys -join ', '))
            foreach ($k in $Query.Keys) {
                $QueryParts += ("{0}={1}" -f (_Encode $k), (_Encode ([string]$Query[$k])))
            }
        }
        if ([string]::IsNullOrWhiteSpace($Search) -eq $false) {
            $HasWild = $Search -match '[\*\?]'
            $CqlOperator = $(if ($HasWild) { "~" } else { "=" })
            $EscapedSearch = $Search.Replace('"','\"')
            $Cql = "title $CqlOperator `"$EscapedSearch`""
            Write-Verbose "Generated CQL: $Cql"
            $QueryParts += ("cql={0}" -f (_Encode $Cql))
        }

        # NEW: Deterministic final endpoint assembly (replaced previous block)
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

        # Final validation (unchanged logic moved below)
        $finalUriObj = $null
        if (-not [System.Uri]::TryCreate($Endpoint, [System.UriKind]::Absolute, [ref]$finalUriObj)) {
            Write-Error "Final endpoint URI invalid. Endpoint='$Endpoint' BaseUriString='$BaseUriString' URIPathNormalized='$URIPathNormalized' QueryPartsCount=$($QueryParts.Count)"
            return
        }
        Write-Verbose "Validated endpoint URI: $($finalUriObj.AbsoluteUri)"

        # --- Init output object ---
        $OutputObjectParams = @{ Results = @(); MultiPage = $false }
        $OutputObject = New-Object -TypeName PSObject -Property $OutputObjectParams
        $script:__ICR_CurrentEndpoint = $Endpoint
    }
    process {
        $QueryPageCount = 0
        $CurrentEndpoint = $script:__ICR_CurrentEndpoint
        Write-Verbose "Starting request loop. MaxQueryPages=$MaxQueryPages Method=$Method"
        do {
            $WRSplat = @{
                Uri         = $CurrentEndpoint
                Headers     = $ConfluenceContext.AuthorizationHeader
                Method      = $Method
                ContentType = "application/json"
            }
            if ($Body) { $WRSplat.Body = $Body }
            Write-Verbose ("[Page {0}] {1} {2}" -f ($QueryPageCount + 1), $Method.ToUpper(), $CurrentEndpoint)
            $Response = Invoke-WebRequest @WRSplat -SkipHttpErrorCheck -ErrorAction Stop

            if ($Response.StatusCode -lt 200 -or $Response.StatusCode -gt 299) {
                $errContent = $null
                try { $errContent = ($Response.Content | ConvertFrom-Json) } catch {}
                $Errors = $errContent.Errors.title
                Write-Error "Failed request. Status: $($Response.StatusCode) $($Response.StatusDescription) Errors: $Errors URL: $CurrentEndpoint"
                return
            }

            $ResponsePayload = $null
            try {
                $ResponsePayload = $Response.Content | ConvertFrom-Json -ErrorAction Stop
            } catch {
                Write-Error "Failed to parse JSON response from $CurrentEndpoint. $_"
                return
            }

            if ($ResponsePayload.PSObject.Properties.Name -contains 'results') {
                $OutputObject.Results += $ResponsePayload.results
                $OutputObject.MultiPage = $true
            } else {
                $OutputObject.Results += $ResponsePayload
            }

            $NextLink = $null
            if ($ResponsePayload.PSObject.Properties.Name -contains 'links') {
                $linksObj = $ResponsePayload.links
                if ($linksObj -and ($linksObj.PSObject.Properties.Name -contains 'next')) {
                    $NextLink = $linksObj.next
                }
            }
            if ($NextLink) {
                if ($NextLink -match '^[a-z]+://') {
                    $CurrentEndpoint = $NextLink
                } else {
                    $CurrentEndpoint = ($BaseUriString.TrimEnd('/')) + ($NextLink.StartsWith('/') ? $NextLink : "/$NextLink")
                }
                Write-Verbose "Detected next page link: $CurrentEndpoint"
            } else {
                $CurrentEndpoint = $null
                Write-Verbose "No further pagination link found."
            }

            $QueryPageCount++
            if ($QueryPageCount -ge $MaxQueryPages) {
                Write-Verbose "Maximum query page count reached. Exiting loop."
                break
            }
        } while ($null -ne $CurrentEndpoint -and $Method -eq 'GET')
        Write-Verbose ("Completed request. Pages retrieved: {0}. Total result objects collected: {1}" -f $QueryPageCount, $OutputObject.Results.Count)
        return $OutputObject
    }
}
