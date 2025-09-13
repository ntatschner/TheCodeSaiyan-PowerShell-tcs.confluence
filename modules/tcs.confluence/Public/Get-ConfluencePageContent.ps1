function Get-ConfluencePageContent {
    [CmdletBinding(DefaultParameterSetName = 'ById')]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ById')]
        [string]$PageId,

        [Parameter(HelpMessage = "Optional space key or numeric spaceId filter (auto-resolved by Invoke-ConfluenceRequest).", ParameterSetName = 'Search')]
        [string]$SpaceKey,

        [Parameter(HelpMessage = "Exact title or wildcard (* ?) to search for.", ParameterSetName = 'Search')]
        [string]$Title,

        [Parameter(HelpMessage = "Maximum results (limit).")]
        [int]$ResultsLimit = 25,

        [Parameter(HelpMessage = "Maximum paged queries.")]
        [int16]$MaxQueryPages = 3,

        [Parameter(Mandatory, HelpMessage = "Content (body) format to return.")]
        [ValidateSet("export_view","storage","editor","view","styled_view","anonymous_export_view","atlas_doc_format")]
        [string]$ContentType = "storage"
    )

    begin {
        $query = @{ 'body-format' = $ContentType }
        if ($PSCmdlet.ParameterSetName -eq 'Search') {
            if ($ResultsLimit) { $query.limit = $ResultsLimit }
            if ($SpaceKey) { $query.spaceKey = $SpaceKey }
            if ($Title -and $Title -notmatch '[\*\?]') {
                # exact title (v2)
                $query.title = ($Title -replace '[<>#%{}|\\^~\[\]`&]', '')
            }
        }
    }

    process {
        try {
            if ($PSCmdlet.ParameterSetName -eq 'ById') {
                $resp = Invoke-ConfluenceRequest -Method GET -Resource pages -Id $PageId -Query $query -MaxQueryPages 1 -ErrorAction Stop
            } else {
                $splat = @{
                    Method        = 'GET'
                    Resource      = 'pages'
                    Query         = $query
                    MaxQueryPages = $MaxQueryPages
                }
                if ($Title -and $Title -match '[\*\?]') { $splat.Search = $Title }
                $resp = Invoke-ConfluenceRequest @splat -ErrorAction Stop
            }
            return $resp
        }
        catch {
            Write-Error "Failed to retrieve page content. Error: $_"
        }
    }
}
