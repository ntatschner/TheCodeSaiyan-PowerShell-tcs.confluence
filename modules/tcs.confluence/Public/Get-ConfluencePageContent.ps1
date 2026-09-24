function Get-ConfluencePageContent {
    <#
    .SYNOPSIS
        Gets Confluence pages including their body in the requested format.

    .DESCRIPTION
        Get-ConfluencePageContent retrieves a page by ID, or searches pages by space and title, and asks
        Confluence to include the page body in the format given by -ContentType (storage by default).
        A title containing wildcards (* ?) is searched with CQL through the v1 content API.

        Returns the object from Invoke-ConfluenceRequest: the pages are in its Results property and
        the body is in body.<format>.value.

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

    .PARAMETER ContentType
        The body format to return: storage (default), atlas_doc_format, view, export_view,
        styled_view, anonymous_export_view or editor.

    .EXAMPLE
        (Get-ConfluencePageContent -PageId 123456).Results.body.storage.value

        Returns the storage-format body of page 123456.

    .EXAMPLE
        Get-ConfluencePageContent -SpaceKey DOCS -Title 'Runbook' -ContentType view

        Returns the Runbook page in the DOCS space with its rendered HTML body.

    .OUTPUTS
        System.Management.Automation.PSCustomObject with Results and MultiPage properties.
    #>
    [CmdletBinding(DefaultParameterSetName = 'ById')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, ParameterSetName = 'ById', HelpMessage = 'The ID of the page.')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Optional space key or numeric spaceId filter (auto-resolved by Invoke-ConfluenceRequest).', ParameterSetName = 'Search')]
        [string]$SpaceKey,

        [Parameter(HelpMessage = 'Exact title or wildcard (* ?) to search for.', ParameterSetName = 'Search')]
        [string]$Title,

        [Parameter(HelpMessage = 'Maximum results (limit).')]
        [ValidateRange(1, 250)]
        [int]$ResultsLimit = 25,

        [Parameter(HelpMessage = 'Maximum paged queries.')]
        [ValidateRange(1, 1000)]
        [int16]$MaxQueryPages = 3,

        [Parameter(HelpMessage = 'Content (body) format to return.')]
        [ValidateSet('export_view', 'storage', 'editor', 'view', 'styled_view', 'anonymous_export_view', 'atlas_doc_format')]
        [string]$ContentType = 'storage'
    )

    $query = @{ 'body-format' = $ContentType }
    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            return (Invoke-ConfluenceRequest -Method GET -Resource pages -Id $PageId -Query $query -MaxQueryPages 1 -ErrorAction Stop)
        }

        if ($ResultsLimit) { $query.limit = $ResultsLimit }
        if ($SpaceKey) { $query.spaceKey = $SpaceKey }
        $requestParams = @{
            Method        = 'GET'
            Resource      = 'pages'
            Query         = $query
            MaxQueryPages = $MaxQueryPages
        }
        if ($Title -and $Title -match '[\*\?]') {
            $requestParams.Search = $Title
        }
        elseif ($Title) {
            # Exact title filter (v2); characters that break the query string are removed
            $query.title = ($Title -replace '[<>#%{}|\\^~\[\]`&]', '')
        }
        return (Invoke-ConfluenceRequest @requestParams -ErrorAction Stop)
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
