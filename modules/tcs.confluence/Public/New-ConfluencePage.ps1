function New-ConfluencePage {
    <#
    .SYNOPSIS
        Creates a Confluence page, or with -Force updates the existing page with the same title.

    .DESCRIPTION
        New-ConfluencePage creates a page under a parent page through the Confluence v2 pages API. The
        body is sent in storage format.

        When Confluence reports that a page with the same title already exists and -Force is given,
        the existing page with exactly that title in the same space (preferring the one under
        -ParentId) is updated with the new content as a new version. A page in another space is never
        updated, and when several pages could match an error is reported instead of choosing one.
        Without -Force the conflict is reported as an error.

        If Confluence reports any other error but a page with exactly this title under this parent
        exists afterwards (Confluence sometimes creates the page and still returns an error), that page
        is returned with a warning.

        Supports -WhatIf and -Confirm. Returns the created or updated page.

    .PARAMETER SpaceId
        The numeric ID of the space to create the page in (the v2 API field spaceId). A space key such
        as DOCS is also accepted and resolved to the ID. -SpaceKey is an alias of this parameter
        (the name used before 0.2.0).

    .PARAMETER ParentId
        The ID of the parent page.

    .PARAMETER Title
        The page title.

    .PARAMETER Status
        The page status: current (published) or draft.

    .PARAMETER Content
        The page body in Confluence storage format (XHTML), for example built with the
        New-ConfluenceContent* functions.

    .PARAMETER Force
        When a page with the same title exists, update it instead of failing. This replaces the
        content of that page.

    .EXAMPLE
        New-ConfluencePage -SpaceId 98765 -ParentId 1000 -Title 'Release 1.2' -Status current -Content '<p>Notes</p>'

        Creates the page "Release 1.2" under page 1000.

    .EXAMPLE
        New-ConfluencePage -SpaceId DOCS -ParentId 1000 -Title 'Daily report' -Status current -Content $html -Force

        Creates the page in the DOCS space, or replaces the content of the existing "Daily report" page.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. The created or updated page.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'The space ID (or a space key, which is resolved to the ID).')]
        [Alias('SpaceKey')]
        [string]$SpaceId,

        [Parameter(Mandatory = $true, HelpMessage = 'The ID of the parent page.')]
        [string]$ParentId,

        [Parameter(Mandatory = $true, HelpMessage = 'The page title.')]
        [string]$Title,

        [Parameter(Mandatory = $true, HelpMessage = 'current or draft.')]
        [ValidateSet('current', 'draft')]
        [string]$Status,

        [Parameter(Mandatory = $true, HelpMessage = 'The page body in storage format.')]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(HelpMessage = "Force overwrite of existing page.`n Destroy existing content.")]
        [switch]$Force
    )

    begin {
        $TelemetryArgs = @{
            ModuleName    = $MyInvocation.MyCommand.Module.Name
            ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
            CommandName   = $MyInvocation.MyCommand.Name
            ExecutionID   = [guid]::NewGuid().ToString()
        }
        Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
        $telemetryFailed = $false
    }

    process {
        try {
            Write-Verbose "Creating Confluence page '$Title' in space '$SpaceId' with parent ID '$ParentId'"
            if (-not $PSCmdlet.ShouldProcess("Confluence page '$Title' in space $SpaceId", 'Create')) {
                return
            }
            try {
                $resolvedSpaceId = Resolve-ConfluenceSpaceId -Space $SpaceId -ErrorAction Stop
            }
            catch {
                Write-Error "Failed to create page '$Title': $($_.Exception.Message)"
                return
            }

            $body = @{
                spaceId  = $resolvedSpaceId
                status   = $Status
                title    = $Title
                parentId = $ParentId
                body     = @{
                    value          = $Content
                    representation = 'storage'
                }
            }

            try {
                $response = Invoke-ConfluenceRequest -Method POST -Resource pages -ApiVersion 2 -Body ($body | ConvertTo-Json -Depth 10) -MaxQueryPages 1 -ErrorAction Stop
                $page = $response.Results | Select-Object -First 1
                Write-Verbose "Created page with ID: $($page.id)"
                return $page
            }
            catch {
                $createError = $_
            }

            $alreadyExists = "$createError" -match '(?i)already exists'
            if ($alreadyExists -and -not $Force) {
                Write-Error "Failed to create page '$Title': a page with this title already exists. Use -Force to update it. $createError"
                return
            }

            # Look for the page with exactly this title (either to update it, or because Confluence may
            # have created it despite reporting an error)
            $existingPage = $null
            try {
                $existingPage = Find-ConfluencePageByTitle -SpaceId $resolvedSpaceId -Title $Title -ParentId $ParentId -ErrorAction Stop
            }
            catch {
                $lookupError = $_
                Write-Verbose "Page lookup after the failed create did not succeed: $lookupError"
                if ($alreadyExists) {
                    Write-Error "Confluence reports that page '$Title' already exists, but it could not be identified safely: $lookupError"
                    return
                }
            }

            if (-not $alreadyExists) {
                if ($existingPage -and "$($existingPage.parentId)" -eq $ParentId) {
                    Write-Warning "Confluence reported an error, but page '$Title' (ID $($existingPage.id)) exists under parent $ParentId; returning it. Error: $createError"
                    return $existingPage
                }
                Write-Error "Failed to create page '$Title'. Error: $createError"
                return
            }

            if (-not $existingPage -or -not $existingPage.id) {
                Write-Error "Confluence reports that page '$Title' already exists, but no page with exactly that title was found in space '$resolvedSpaceId'. Error: $createError"
                return
            }

            $newVersion = 1
            if ($existingPage.version -and $existingPage.version.number) {
                $newVersion = [int]$existingPage.version.number + 1
            }
            Write-Verbose "Page '$Title' already exists (ID $($existingPage.id)); updating it to version $newVersion."
            try {
                return (Update-ConfluencePage -PageId $existingPage.id -Title $Title -Status $Status -Content $Content -Version $newVersion -ErrorAction Stop)
            }
            catch {
                Write-Error "Failed to update existing page '$Title' (ID $($existingPage.id)). Error: $_"
            }
        }
        catch {
            if (-not $telemetryFailed) {
                $telemetryFailed = $true
                Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
            }
            throw
        }
    }

    end {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
