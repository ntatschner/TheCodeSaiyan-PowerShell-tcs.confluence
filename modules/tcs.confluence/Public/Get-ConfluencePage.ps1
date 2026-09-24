function Get-ConfluencePage {
    <#
    .SYNOPSIS
        Gets Confluence pages by ID, by space or by title.

    .DESCRIPTION
        Get-ConfluencePage retrieves pages through the Confluence v2 pages API. Use -PageId for a single
        page, or -SpaceKey and/or -Search to list pages. An exact -Search title is sent as a title
        filter; a title with wildcards (* ?) is sent as a CQL search through the v1 content API.

        Returns the object from Invoke-ConfluenceRequest: the pages are in its Results property.

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

    .EXAMPLE
        (Get-ConfluencePage -PageId 123456).Results

        Returns page 123456.

    .EXAMPLE
        (Get-ConfluencePage -SpaceKey DOCS -Search 'Release notes*').Results | Select-Object id, title

        Lists the pages in the DOCS space whose title starts with "Release notes".

    .OUTPUTS
        System.Management.Automation.PSCustomObject with Results and MultiPage properties.
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
        [int16]$MaxQueryPages = 3
    )

    $queryParams = @{}
    $requestParams = @{
        Method        = 'GET'
        Resource      = 'pages'
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
    }
    $requestParams.Query = $queryParams
    if ($queryParams.Count -gt 0) {
        Write-Verbose ('Query parameters: {0}' -f (($queryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '))
    }

    try {
        $response = Invoke-ConfluenceRequest @requestParams -ErrorAction Stop
        if (-not $response.Results -or @($response.Results).Count -eq 0) { Write-Verbose 'No results returned.' }
        return $response
    }
    catch {
        Write-Error "Failed to retrieve page information. Error: $_"
    }
}
