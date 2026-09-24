function Update-ConfluencePage {
    <#
    .SYNOPSIS
        Replaces the title and body of a Confluence page.

    .DESCRIPTION
        Update-ConfluencePage sends a new version of a page through the Confluence v2 pages API. The
        body is sent in storage format. Confluence requires the new version number, which is the
        current version number plus one.

        Supports -WhatIf and -Confirm. Returns the updated page.

    .PARAMETER PageId
        The ID of the page to update.

    .PARAMETER SpaceKey
        The numeric space ID to move the page to. Leave empty to keep the page in its space.

    .PARAMETER Title
        The page title.

    .PARAMETER Status
        The page status: current (default) or draft.

    .PARAMETER Content
        The page body in Confluence storage format (XHTML).

    .PARAMETER Version
        The new version number: the page's current version number plus one.

    .PARAMETER VersionMessage
        The version comment shown in the page history. Default "Programmatically Updated".

    .EXAMPLE
        $page = (Get-ConfluencePage -PageId 123456).Results[0]
        Update-ConfluencePage -PageId $page.id -Title $page.title -Content '<p>New body</p>' -Version ($page.version.number + 1)

        Replaces the body of page 123456.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. The updated page.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The page ID of the page to update.')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Space key / id (optional if unchanged).')]
        [string]$SpaceKey,

        [Parameter(Mandatory, HelpMessage = 'Title of the page.')]
        [string]$Title,

        [Parameter(HelpMessage = 'Status: draft or current.')]
        [ValidateSet('draft', 'current')]
        [string]$Status = 'current',

        [Parameter(Mandatory, HelpMessage = 'Storage format content.')]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(Mandatory, HelpMessage = 'New version number (increment previous).')]
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
            $body = @{
                id      = $PageId
                title   = $Title
                status  = $Status
                body    = @{
                    value          = $Content
                    representation = 'storage'
                }
                version = @{
                    number  = $Version
                    message = $VersionMessage
                }
            }
            if ($SpaceKey) { $body.spaceId = $SpaceKey }

            if (-not $PSCmdlet.ShouldProcess("Confluence page $PageId ('$Title')", "Update to version $Version")) {
                return
            }

            try {
                $response = Invoke-ConfluenceRequest -Method PUT -Resource pages -Id $PageId -Body ($body | ConvertTo-Json -Depth 10) -MaxQueryPages 1 -ErrorAction Stop
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
