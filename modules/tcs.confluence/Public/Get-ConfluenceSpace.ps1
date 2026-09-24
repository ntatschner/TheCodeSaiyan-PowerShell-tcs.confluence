function Get-ConfluenceSpace {
    <#
    .SYNOPSIS
        Gets Confluence spaces, optionally filtered by name.

    .DESCRIPTION
        Get-ConfluenceSpace lists the spaces visible to the account in the Confluence context, or
        returns one space by ID. -Search filters the listed spaces by name on the client side and
        supports wildcards (* ?).

        Called as Get-ConfluenceSpaces (the name used before 0.1.0) it still works through an alias.

    .PARAMETER SpaceId
        The numeric ID of the space to retrieve. Returns the object from Invoke-ConfluenceRequest.

    .PARAMETER Search
        A space name filter, for example 'Team*'. Matching is case-insensitive.

    .PARAMETER ResultsLimit
        The number of spaces requested per API call. Default 50.

    .PARAMETER MaxQueryPages
        The maximum number of API result pages to retrieve. Default 3.

    .EXAMPLE
        Get-ConfluenceSpace | Select-Object id, key, name

        Lists the spaces.

    .EXAMPLE
        Get-ConfluenceSpace -Search 'Engineering*'

        Lists the spaces whose name starts with "Engineering".

    .OUTPUTS
        System.Management.Automation.PSCustomObject
    #>
    [CmdletBinding(DefaultParameterSetName = 'AllSpaces')]
    [Alias('Get-ConfluenceSpaces')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'SpaceId', HelpMessage = 'The ID of the space.')]
        [string]$SpaceId,

        [Parameter(ParameterSetName = 'AllSpaces', HelpMessage = 'Name filter (supports wildcard * ?) applied client-side.')]
        [string]$Search,

        [Parameter(ParameterSetName = 'AllSpaces', HelpMessage = 'Results per request.')]
        [ValidateRange(1, 250)]
        [int]$ResultsLimit = 50,

        [Parameter(ParameterSetName = 'AllSpaces', HelpMessage = 'Maximum number of result pages.')]
        [ValidateRange(1, 1000)]
        [int16]$MaxQueryPages = 3
    )

    try {
        if ($PSCmdlet.ParameterSetName -eq 'SpaceId') {
            return (Invoke-ConfluenceRequest -Method GET -Resource spaces -Id $SpaceId -MaxQueryPages 1 -ErrorAction Stop)
        }

        $response = Invoke-ConfluenceRequest -Method GET -Resource spaces -Query @{ limit = $ResultsLimit } -MaxQueryPages $MaxQueryPages -ErrorAction Stop
        $results = $response.Results
        if ($Search) {
            $results = $results | Where-Object { $_.name -like $Search }
        }
        return $results
    }
    catch {
        Write-Error "Failed to retrieve space information. Error: $_"
    }
}
