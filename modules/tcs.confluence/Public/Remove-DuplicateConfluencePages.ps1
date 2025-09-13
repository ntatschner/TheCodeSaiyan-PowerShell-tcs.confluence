function Remove-DuplicateConfluencePages {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory)]
        [string]$SpaceKey,
        [Parameter(Mandatory)]
        [string]$ParentId,
        [Parameter(Mandatory)]
        [string]$PageTitle,
        [Parameter()]
        [switch]$KeepNewest
    )

    begin {
        if (-not (Get-Variable -Scope Global -Name "ConfluenceContext" -ErrorAction SilentlyContinue)) {
            Write-Error "ConfluenceContext variable not found. Run Set-ConfluenceContext first."
            return
        }
        if (-not $PSBoundParameters.ContainsKey('KeepNewest')) { $KeepNewest = $true }
        Write-Verbose "Looking for duplicate pages Title='$PageTitle' ParentId='$ParentId' SpaceKey='$SpaceKey' KeepNewest=$KeepNewest"
    }

    process {
        try {
            $all = Get-ConfluencePage -SpaceKey $SpaceKey -ErrorAction Stop
            $pagesCollection = if ($all.PSObject.Properties.Name -contains 'results') { $all.results } elseif ($all.PSObject.Properties.Name -contains 'Results') { $all.Results } else { @() }

            if (-not $pagesCollection) {
                Write-Verbose "No pages returned for space '$SpaceKey'."
                return $null
            }

            $duplicates = $pagesCollection | Where-Object { $_.title -eq $PageTitle -and $_.parentId -eq $ParentId }

            if (-not $duplicates -or $duplicates.Count -le 1) {
                Write-Verbose "Zero or one page found with that title/parent. Nothing to remove."
                return ($duplicates | Select-Object -First 1)
            }

            Write-Verbose "Found $($duplicates.Count) pages matching criteria."

            $sortProperties = @(
                @{ Expression = { try { [int]$_.version.number } catch { 0 } } },
                @{ Expression = { try { [datetime]$_.lastModified } catch { [datetime]::MinValue } } }
            )
            $sortedPages = $duplicates | Sort-Object -Property $sortProperties -Descending:$KeepNewest

            $pageToKeep = $sortedPages | Select-Object -First 1
            $pagesToDelete = $sortedPages | Select-Object -Skip 1

            $versionNum = if ($pageToKeep.version) { $pageToKeep.version.number } else { 'N/A' }
            Write-Verbose ("Keeping page ID={0} Version={1}" -f $pageToKeep.id, $versionNum)

            foreach ($p in $pagesToDelete) {
                $deleteId = $p.id
                if (-not $PSCmdlet.ShouldProcess("page ID $deleteId", "Delete")) {
                    Write-Verbose "Skipping deletion of page ID=$deleteId due to -WhatIf or user choice."
                    continue
                }
                try {
                    Write-Verbose "Deleting page ID=$deleteId"
                    $null = Invoke-ConfluenceRequest -Method DELETE -Resource pages -Id $deleteId -MaxQueryPages 1 -ErrorAction Stop
                    Write-Verbose "Deleted page ID=$deleteId"
                } catch {
                    Write-Warning ("Failed to delete page ID {0}. Error: {1}" -f $deleteId, $_)
                }
            }

            return $pageToKeep
        }
        catch {
            Write-Error "Error removing duplicate pages. $_"
        }
    }
}
