function Remove-DuplicateConfluencePage {
    <#
    .SYNOPSIS
        Removes duplicate Confluence pages that share a title and parent, keeping one.

    .DESCRIPTION
        Remove-DuplicateConfluencePage lists the pages in a space, finds the pages with the given title
        under the given parent and, when there is more than one, deletes all but one. By default the
        newest page (highest version number, then latest modification) is kept; use
        -KeepNewest:$false to keep the oldest.

        Every deletion asks for confirmation unless you pass -Confirm:$false; -WhatIf shows what would
        be deleted. The kept page is returned. Only the pages returned by Get-ConfluencePage for the
        space (up to three API result pages) are considered.

        Called as Remove-DuplicateConfluencePages (the name used before 0.1.0) it still works through
        an alias.

    .PARAMETER SpaceKey
        The key or numeric ID of the space to search.

    .PARAMETER ParentId
        The ID of the parent page of the duplicates.

    .PARAMETER PageTitle
        The exact title of the duplicated page.

    .PARAMETER KeepNewest
        Keep the newest page (default). Use -KeepNewest:$false to keep the oldest page instead.

    .EXAMPLE
        Remove-DuplicateConfluencePage -SpaceKey DOCS -ParentId 1000 -PageTitle 'Weekly report' -WhatIf

        Shows which duplicate "Weekly report" pages under page 1000 would be deleted.

    .EXAMPLE
        Remove-DuplicateConfluencePage -SpaceKey DOCS -ParentId 1000 -PageTitle 'Weekly report' -KeepNewest:$false -Confirm:$false

        Deletes all but the oldest duplicate without prompting.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. The page that was kept.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [Alias('Remove-DuplicateConfluencePages')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The space key or ID.')]
        [string]$SpaceKey,

        [Parameter(Mandatory, HelpMessage = 'The ID of the parent page.')]
        [string]$ParentId,

        [Parameter(Mandatory, HelpMessage = 'The exact page title.')]
        [string]$PageTitle,

        [Parameter(HelpMessage = 'Keep the newest page (default) or, with -KeepNewest:$false, the oldest.')]
        [switch]$KeepNewest
    )

    if (-not $PSBoundParameters.ContainsKey('KeepNewest')) { $KeepNewest = $true }
    Write-Verbose "Looking for duplicate pages Title='$PageTitle' ParentId='$ParentId' SpaceKey='$SpaceKey' KeepNewest=$KeepNewest"

    try {
        $all = Get-ConfluencePage -SpaceKey $SpaceKey -ErrorAction Stop
        $pagesCollection = @($all.Results | Where-Object { $null -ne $_ })
        if ($pagesCollection.Count -eq 0) {
            Write-Verbose "No pages returned for space '$SpaceKey'."
            return
        }

        $duplicates = @($pagesCollection | Where-Object { $_.title -eq $PageTitle -and "$($_.parentId)" -eq $ParentId })
        if ($duplicates.Count -le 1) {
            Write-Verbose 'Zero or one page found with that title/parent. Nothing to remove.'
            return ($duplicates | Select-Object -First 1)
        }
        Write-Verbose "Found $($duplicates.Count) pages matching criteria."

        $sortProperties = @(
            @{ Expression = { $number = 0; if ($_.version -and [int]::TryParse("$($_.version.number)", [ref]$number)) { $number } else { 0 } } },
            @{ Expression = { $date = [datetime]::MinValue; $stamp = if ($_.version -and $_.version.createdAt) { $_.version.createdAt } else { $_.lastModified }; if ($stamp -and [datetime]::TryParse("$stamp", [ref]$date)) { $date } else { [datetime]::MinValue } } }
        )
        $sortedPages = @($duplicates | Sort-Object -Property $sortProperties -Descending:$KeepNewest)
        $pageToKeep = $sortedPages[0]
        $pagesToDelete = @($sortedPages | Select-Object -Skip 1)

        $versionNum = if ($pageToKeep.version) { $pageToKeep.version.number } else { 'N/A' }
        Write-Verbose ('Keeping page ID={0} Version={1}' -f $pageToKeep.id, $versionNum)

        foreach ($page in $pagesToDelete) {
            $deleteId = $page.id
            if (-not $PSCmdlet.ShouldProcess("Confluence page $deleteId ('$PageTitle')", 'Remove duplicate')) {
                continue
            }
            try {
                $null = Invoke-ConfluenceRequest -Method DELETE -Resource pages -Id $deleteId -MaxQueryPages 1 -ErrorAction Stop
                Write-Verbose "Deleted page ID=$deleteId"
            }
            catch {
                Write-Error ('Failed to delete page ID {0}. Error: {1}' -f $deleteId, $_)
            }
        }

        return $pageToKeep
    }
    catch {
        Write-Error "Error removing duplicate pages. $_"
    }
}
