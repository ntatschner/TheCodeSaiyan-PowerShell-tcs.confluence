function Update-ConfluencePage {
    <#
    .SYNOPSIS
        Replaces the title and body of a Confluence page.

    .DESCRIPTION
        Update-ConfluencePage sends a new version of a page through the Confluence v2 pages API. The
        body is sent in storage format. Confluence requires the new version number, which is the
        current version number plus one: pass it with -Version, or leave -Version out and the current
        page is read first and its version number plus one is used.

        Supports -WhatIf and -Confirm. Returns the updated page.

    .PARAMETER PageId
        The ID of the page to update.

    .PARAMETER SpaceId
        The numeric space ID (or a space key, resolved to the ID) to move the page to. Leave empty to
        keep the page in its space. -SpaceKey is an alias of this parameter (the name used before
        0.2.0).

    .PARAMETER Title
        The page title.

    .PARAMETER Status
        The page status: current (default) or draft.

    .PARAMETER Content
        The page body in Confluence storage format (XHTML).

    .PARAMETER Version
        The new version number: the page's current version number plus one. When not given, the page
        is read and its current version number plus one is used.

    .PARAMETER VersionMessage
        The version comment shown in the page history. Default "Programmatically Updated".

    .EXAMPLE
        Update-ConfluencePage -PageId 123456 -Title 'Runbook' -Content '<p>New body</p>'

        Replaces the body of page 123456 as its next version.

    .EXAMPLE
        $page = Get-ConfluencePage -PageId 123456
        Update-ConfluencePage -PageId $page.id -Title $page.title -Content '<p>New body</p>' -Version ($page.version.number + 1)

        Replaces the body of page 123456 with an explicit version number.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. The updated page.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The page ID of the page to update.')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Space id or key (optional if unchanged).')]
        [Alias('SpaceKey')]
        [string]$SpaceId,

        [Parameter(Mandatory, HelpMessage = 'Title of the page.')]
        [string]$Title,

        [Parameter(HelpMessage = 'Status: draft or current.')]
        [ValidateSet('draft', 'current')]
        [string]$Status = 'current',

        [Parameter(Mandatory, HelpMessage = 'Storage format content.')]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(HelpMessage = 'New version number (increment previous). Default: current version + 1.')]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Version,

        [Parameter(HelpMessage = 'The version comment.')]
        [string]$VersionMessage = 'Programmatically Updated'
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
            $versionText = if ($Version) { "version $Version" } else { 'the next version' }
            if (-not $PSCmdlet.ShouldProcess("Confluence page $PageId ('$Title')", "Update to $versionText")) {
                return
            }

            $newVersion = $Version
            if (-not $newVersion) {
                try {
                    $current = (Invoke-ConfluenceRequest -Method GET -Resource pages -ApiVersion 2 -Id $PageId -MaxQueryPages 1 -ErrorAction Stop).Results | Select-Object -First 1
                }
                catch {
                    Write-Error "Failed to read the current version of page '$PageId'. Error: $_"
                    return
                }
                if (-not $current -or -not $current.version -or -not $current.version.number) {
                    Write-Error "Could not read the current version number of page '$PageId'. Pass -Version."
                    return
                }
                $newVersion = [int]$current.version.number + 1
                Write-Verbose "Page $PageId is at version $($current.version.number); sending version $newVersion."
            }

            $body = @{
                id      = $PageId
                title   = $Title
                status  = $Status
                body    = @{
                    value          = $Content
                    representation = 'storage'
                }
                version = @{
                    number  = $newVersion
                    message = $VersionMessage
                }
            }
            if ($SpaceId) {
                try {
                    $body.spaceId = Resolve-ConfluenceSpaceId -Space $SpaceId -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed to update page '$PageId': $($_.Exception.Message)"
                    return
                }
            }

            try {
                $response = Invoke-ConfluenceRequest -Method PUT -Resource pages -ApiVersion 2 -Id $PageId -Body ($body | ConvertTo-Json -Depth 10) -MaxQueryPages 1 -ErrorAction Stop
                return ($response.Results | Select-Object -First 1)
            }
            catch {
                Write-Error "Failed to update page '$PageId'. Error: $_"
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
