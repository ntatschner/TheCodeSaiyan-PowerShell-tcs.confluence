function Find-ConfluencePageByTitle {
    <#
    .SYNOPSIS
        Finds a page in a space by its exact title, preferring the page under the given parent.

    .DESCRIPTION
        Used by New-ConfluencePage to resolve "page already exists" conflicts. Only pages whose title
        matches exactly are returned; when several match, the one whose parentId equals ParentId wins.
        Returns nothing when no page matches.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$SpaceKey,

        [Parameter(Mandatory)]
        [string]$Title,

        [string]$ParentId
    )

    $response = Get-ConfluencePage -SpaceKey $SpaceKey -Search $Title -ErrorAction Stop
    $candidates = @($response.Results | Where-Object { $null -ne $_ -and $_.title -eq $Title })
    if ($candidates.Count -eq 0) {
        return
    }
    if ($ParentId) {
        $underParent = @($candidates | Where-Object { "$($_.parentId)" -eq $ParentId })
        if ($underParent.Count -gt 0) {
            return $underParent[0]
        }
    }
    return $candidates[0]
}
