function Get-ConfluencePage {
    <#
    .SYNOPSIS
        Gets Confluence pages by ID, by space or by title.

    .DESCRIPTION
        Get-ConfluencePage retrieves pages through the Confluence v2 pages API and writes the page
        objects to the pipeline. Use -PageId for a single page, or -SpaceKey and/or -Search to list
        pages. An exact -Search title is sent as the v2 title filter; a title with wildcards (* ?) is
        searched with CQL through /wiki/rest/api/content/search (those results are v1 content
        objects).

        Up to -MaxQueryPages API result pages are read; a warning is written when more results are
        available. Use -All to read every result page.

        Before 0.2.0 the pages were wrapped in an object with Results and MultiPage properties.

    .PARAMETER PageId
        The ID of the page to retrieve.

    .PARAMETER SpaceKey
        The key (or numeric ID) of the space whose pages are listed. Keys are resolved to space IDs.

    .PARAMETER Search
        A page title to match. Wildcards (* ?) are supported.

    .PARAMETER ResultsLimit
        The number of results requested per API call. Default 25.

    .PARAMETER MaxQueryPages
        The maximum number of API result pages to retrieve. Default 3.

    .PARAMETER All
        Retrieve every API result page (ignores -MaxQueryPages).

    .EXAMPLE
        Get-ConfluencePage -PageId 123456

        Returns page 123456.

    .EXAMPLE
        Get-ConfluencePage -SpaceKey DOCS -Search 'Release notes*' -All | Select-Object id, title

        Lists every page in the DOCS space whose title starts with "Release notes".

    .EXAMPLE
        Get-ConfluencePage -SpaceKey DOCS -Search 'Draft*' | Remove-ConfluencePage -WhatIf

        Shows which pages would be deleted.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. One object per page.
    #>
    [CmdletBinding(DefaultParameterSetName = 'AllPages')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'The page ID of the page to retrieve.', ParameterSetName = 'PageId')]
        [string]$PageId,

        [Parameter(HelpMessage = 'The space key of the page(s) to retrieve.', ParameterSetName = 'AllPages')]
        [string]$SpaceKey,

        [Parameter(HelpMessage = 'Page title search string. Wildcards (* ?) supported.', ParameterSetName = 'AllPages')]
        [string]$Search,

        [Parameter(HelpMessage = 'Max results per request')]
        [ValidateRange(1, 250)]
        [int]$ResultsLimit = 25,

        [Parameter(HelpMessage = 'Max query pages.')]
        [ValidateRange(1, 1000)]
        [int16]$MaxQueryPages = 3,

        [Parameter(HelpMessage = 'Retrieve every result page.', ParameterSetName = 'AllPages')]
        [switch]$All
    )

    $queryParams = @{}
    $requestParams = @{
        Method        = 'GET'
        Resource      = 'pages'
        ApiVersion    = 2
        MaxQueryPages = $MaxQueryPages
    }
    if ($PSCmdlet.ParameterSetName -eq 'PageId') {
        $requestParams.Id = $PageId
    }
    else {
        $queryParams.limit = $ResultsLimit
        if ($SpaceKey) { $queryParams.spaceKey = $SpaceKey }
        if ($Search) {
            if ($Search -match '[\*\?]') {
                $requestParams.Search = $Search
            }
            else {
                $queryParams.title = $Search
            }
        }
        if ($All) { $requestParams.All = $true }
    }
    $requestParams.Query = $queryParams

    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        $response = Invoke-ConfluenceRequest @requestParams -ErrorAction Stop
        if (@($response.Results).Count -eq 0) { Write-Verbose 'No results returned.' }
        foreach ($page in @($response.Results)) { $page }
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        Write-Error "Failed to retrieve page information. Error: $_"
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
