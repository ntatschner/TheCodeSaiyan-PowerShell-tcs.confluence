function Search-ConfluenceContent {
    <#
    .SYNOPSIS
        Searches Confluence with a CQL query.

    .DESCRIPTION
        Search-ConfluenceContent runs a Confluence Query Language (CQL) search through the v1 search
        API (GET /wiki/rest/api/search?cql=...) and writes each search result to the pipeline. A
        result has content (the page, blog post or attachment), title, excerpt, url and
        lastModified properties.

        Up to -MaxQueryPages result pages are read; a warning is written when more results are
        available. Use -All to read every result page.

        The query is sent as given: quote values with double quotes and escape a backslash or double
        quote inside a value with a backslash.

    .PARAMETER Cql
        The CQL query, for example: type = page AND space = DOCS AND text ~ "backup".

    .PARAMETER Limit
        The number of results requested per API call. Default 25.

    .PARAMETER MaxQueryPages
        The maximum number of API result pages to retrieve. Default 3.

    .PARAMETER All
        Retrieve every API result page (ignores -MaxQueryPages).

    .PARAMETER Expand
        Properties to expand in the results, for example content.space or content.version.

    .EXAMPLE
        Search-ConfluenceContent -Cql 'type = page AND space = DOCS AND lastmodified > now("-7d")' -All

        Returns every page in the DOCS space changed in the last seven days.

    .EXAMPLE
        Search-ConfluenceContent -Cql 'label = runbook' | Select-Object title, url

        Lists the content labelled runbook.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. One object per search result.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, Position = 0, HelpMessage = 'The CQL query.')]
        [ValidateNotNullOrEmpty()]
        [string]$Cql,

        [Parameter(HelpMessage = 'Results per request.')]
        [ValidateRange(1, 1000)]
        [int]$Limit = 25,

        [Parameter(HelpMessage = 'Maximum number of result pages.')]
        [ValidateRange(1, 1000)]
        [int16]$MaxQueryPages = 3,

        [Parameter(HelpMessage = 'Retrieve every result page.')]
        [switch]$All,

        [Parameter(HelpMessage = 'Properties to expand in the results.')]
        [string[]]$Expand
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
        $query = @{ cql = $Cql; limit = $Limit }
        if ($Expand) { $query.expand = $Expand -join ',' }
        $response = Invoke-ConfluenceRequest -Method GET -URIPath '/wiki/rest/api/search' -Query $query -MaxQueryPages $MaxQueryPages -All:$All -ErrorAction Stop
        foreach ($result in @($response.Results)) { $result }
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        Write-Error "Confluence search failed. Error: $_"
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
