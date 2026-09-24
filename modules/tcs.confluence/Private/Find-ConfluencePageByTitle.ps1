function Find-ConfluencePageByTitle {
    <#
    .SYNOPSIS
        Finds the page in a space with an exact title, preferring the page under the given parent.

    .DESCRIPTION
        Used by New-ConfluencePage to resolve "page already exists" conflicts. Only pages whose title
        matches exactly and whose spaceId is the requested space are considered, so a page in another
        space is never returned. When one page is under ParentId it is returned; otherwise the single
        matching page in the space is returned. When several pages match and none (or several) is
        under ParentId, an error is thrown instead of guessing. Returns nothing when no page matches.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$SpaceId,

        [Parameter(Mandatory)]
        [string]$Title,

        [string]$ParentId
    )

    $resolvedSpaceId = Resolve-ConfluenceSpaceId -Space $SpaceId -ErrorAction Stop
    $response = Invoke-ConfluenceRequest -Method GET -Resource pages -ApiVersion 2 -Query @{ 'space-id' = $resolvedSpaceId; title = $Title; limit = 250 } -All -ErrorAction Stop
    $candidates = @($response.Results | Where-Object {
            $null -ne $_ -and $_.title -eq $Title -and "$($_.spaceId)" -eq $resolvedSpaceId
        })
    if ($candidates.Count -eq 0) {
        return
    }
    if ($ParentId) {
        $underParent = @($candidates | Where-Object { "$($_.parentId)" -eq $ParentId })
        if ($underParent.Count -eq 1) {
            return $underParent[0]
        }
        if ($underParent.Count -gt 1) {
            throw "Found $($underParent.Count) pages titled '$Title' under parent $ParentId in space $resolvedSpaceId (IDs $((@($underParent.id)) -join ', ')); not choosing one."
        }
    }
    if ($candidates.Count -eq 1) {
        return $candidates[0]
    }
    throw "Found $($candidates.Count) pages titled '$Title' in space $resolvedSpaceId (IDs $((@($candidates.id)) -join ', ')); not choosing one."
}
