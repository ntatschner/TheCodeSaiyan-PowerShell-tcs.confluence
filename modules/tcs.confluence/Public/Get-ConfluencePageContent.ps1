function Get-ConfluencePageContent {
    <#
    .SYNOPSIS
        Gets Confluence pages including their body in the requested format.

    .DESCRIPTION
        Get-ConfluencePageContent retrieves a page by ID, or searches pages by space and title, and asks
        Confluence to include the page body in the format given by -ContentType (storage by default)
        and writes the page objects to the pipeline. The body is in body.<format>.value.

        By ID and exact titles use the v2 pages API (body-format=<format>). A title containing
        wildcards (* ?) is searched with CQL through /wiki/rest/api/content/search, which returns v1
        content objects; the body is requested with expand=body.<format>.

        Up to -MaxQueryPages API result pages are read; a warning is written when more results are
        available. Use -All to read every result page.

        Before 0.2.0 the pages were wrapped in an object with Results and MultiPage properties.

    .PARAMETER PageId
        The ID of the page to retrieve.

    .PARAMETER SpaceKey
        The space key or numeric space ID to search in.

    .PARAMETER Title
        The exact page title, or a title with wildcards (* ?).

    .PARAMETER ResultsLimit
        The number of results requested per API call when searching. Default 25.

    .PARAMETER MaxQueryPages
        The maximum number of API result pages to retrieve when searching. Default 3.

    .PARAMETER All
        Retrieve every API result page when searching (ignores -MaxQueryPages).

    .PARAMETER ContentType
        The body format to return: storage (default), atlas_doc_format, view, export_view,
        styled_view, anonymous_export_view or editor.

    .EXAMPLE
        (Get-ConfluencePageContent -PageId 123456).body.storage.value

        Returns the storage-format body of page 123456.

    .EXAMPLE
        Get-ConfluencePageContent -SpaceKey DOCS -Title 'R&D [draft]' -ContentType view

        Returns the "R&D [draft]" page in the DOCS space with its rendered HTML body.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. One object per page.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ById')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ById', HelpMessage = 'The ID of the page.')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Optional space key or numeric spaceId filter.', ParameterSetName = 'Search')]
        [string]$SpaceKey,

        [Parameter(HelpMessage = 'Exact title or wildcard (* ?) to search for.', ParameterSetName = 'Search')]
        [string]$Title,

        [Parameter(HelpMessage = 'Maximum results (limit).')]
        [ValidateRange(1, 250)]
        [int]$ResultsLimit = 25,

        [Parameter(HelpMessage = 'Maximum paged queries.')]
        [ValidateRange(1, 1000)]
        [int16]$MaxQueryPages = 3,

        [Parameter(HelpMessage = 'Retrieve every result page.', ParameterSetName = 'Search')]
        [switch]$All,

        [Parameter(HelpMessage = 'Content (body) format to return.')]
        [ValidateSet('export_view', 'storage', 'editor', 'view', 'styled_view', 'anonymous_export_view', 'atlas_doc_format')]
        [string]$ContentType = 'storage'
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
        $requestParams = @{
            Method        = 'GET'
            Resource      = 'pages'
            ApiVersion    = 2
            MaxQueryPages = $MaxQueryPages
        }
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            $requestParams.Id = $PageId
            $requestParams.MaxQueryPages = 1
            $requestParams.Query = @{ 'body-format' = $ContentType }
        }
        else {
            $query = @{ limit = $ResultsLimit }
            if ($SpaceKey) { $query.spaceKey = $SpaceKey }
            if ($Title -and $Title -match '[\*\?]') {
                # CQL search (v1 content objects): the body is requested with expand
                $requestParams.Search = $Title
                $query.expand = "body.$ContentType,version,space"
            }
            else {
                $query['body-format'] = $ContentType
                # Exact title filter (v2); the value is URL-encoded, so it is sent unchanged
                if ($Title) { $query.title = $Title }
            }
            $requestParams.Query = $query
            if ($All) { $requestParams.All = $true }
        }

        $response = Invoke-ConfluenceRequest @requestParams -ErrorAction Stop
        foreach ($page in @($response.Results)) { $page }
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        Write-Error "Failed to retrieve page content. Error: $_"
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
