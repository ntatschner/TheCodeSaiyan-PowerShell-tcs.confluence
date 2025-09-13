function Get-ConfluencePage {
    [CmdletBinding(DefaultParameterSetName = 'AllPages')]
    param (
        [Parameter(HelpMessage = "The page ID of the page to retrieve.",
            ParameterSetName = 'PageId')]
        [string]$PageId,

        [Parameter(HelpMessage = "The space key of the page(s) to retrieve.",
            ParameterSetName = 'AllPages')]
        [string]$SpaceKey,

        [Parameter(HelpMessage = "Page title search string. Wildcards (* ?) supported.",
            ParameterSetName = 'AllPages')]
        [string]$Search,

        [Parameter(HelpMessage = "Max results per request")]
        [int]$ResultsLimit = 25,

        [Parameter(HelpMessage = "Max query pages.")]
        [int16]$MaxQueryPages = 3
    )

    begin {
        $QueryParams = @{ limit = $ResultsLimit }
        if ($PSCmdlet.ParameterSetName -eq 'AllPages') {
            if ($SpaceKey) { $QueryParams.spaceKey = $SpaceKey }
            if ($Search -and $Search -notmatch '[\*\?]') {
                $QueryParams.title = $Search
            }
        }
        $script:__GCP_Request = @{
            Query         = $QueryParams
            Search        = $Search
            PageId        = $PageId
            MaxQueryPages = $MaxQueryPages
            ParamSet      = $PSCmdlet.ParameterSetName
        }
    }
    process {
        try {
            $r = $script:__GCP_Request
            $splat = @{
                Method        = 'GET'
                Resource      = 'pages'
                Query         = $r.Query
                MaxQueryPages = $r.MaxQueryPages
            }
            if ($r.ParamSet -eq 'PageId') { $splat.Id = $r.PageId }
            elseif ($r.Search -and $r.Search -match '[\*\?]') { $splat.Search = $r.Search }
            if ($r.Query.Count -gt 0) {
                Write-Verbose ("Query parameters: {0}" -f (($r.Query.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '))
            }
            Write-Verbose "Invoking Get-ConfluencePage via -Resource pages"
            $resp = Invoke-ConfluenceRequest @splat -ErrorAction Stop
            if (-not $resp.Results -or $resp.Results.Count -eq 0) { Write-Verbose "No results returned." }
            return $resp
        } catch {
            Write-Error "Failed to retrieve page information. Error: $_"
        }
    }
}
