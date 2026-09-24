function Remove-ConfluencePage {
    <#
    .SYNOPSIS
        Deletes a Confluence page.

    .DESCRIPTION
        Remove-ConfluencePage deletes the page with the given ID through the Confluence v2 pages API.
        The page moves to the space trash; with -Purge it is then permanently deleted from the trash
        (this needs space admin permission and cannot be undone). Because this changes the site, you
        are asked to confirm unless you pass -Confirm:$false; -WhatIf shows what would be deleted.

    .PARAMETER PageId
        The ID of the page to delete. Accepts pipeline input by property name (id), for example the
        pages written by Get-ConfluencePage.

    .PARAMETER Purge
        Permanently delete the page: it is moved to the trash (if it is not there already) and then
        purged with DELETE /wiki/api/v2/pages/<id>?purge=true.

    .EXAMPLE
        Remove-ConfluencePage -PageId 123456

        Asks for confirmation, then deletes page 123456.

    .EXAMPLE
        Get-ConfluencePage -SpaceKey DOCS -Search 'Draft*' | Remove-ConfluencePage -WhatIf

        Shows which draft pages would be deleted without deleting them.

    .EXAMPLE
        Remove-ConfluencePage -PageId 123456 -Purge -Confirm:$false

        Deletes page 123456 and purges it from the trash.

    .OUTPUTS
        None.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'The ID of the page to remove.')]
        [Alias('id')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Permanently delete the page (trash, then purge).')]
        [switch]$Purge
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
            $action = if ($Purge) { 'Remove and purge' } else { 'Remove' }
            if ($PSCmdlet.ShouldProcess("Confluence page $PageId", $action)) {
                $trashError = $null
                try {
                    $null = Invoke-ConfluenceRequest -Method DELETE -Resource pages -ApiVersion 2 -Id $PageId -MaxQueryPages 1 -ErrorAction Stop
                    Write-Verbose "Moved page $PageId to the trash."
                }
                catch {
                    $trashError = $_
                    if (-not $Purge) {
                        Write-Error "Failed to remove page with ID $PageId. Error: $trashError"
                    }
                }
                if ($Purge) {
                    # Purging only works on a trashed page; if moving it to the trash failed, the page
                    # may already be in the trash, so the purge is still attempted
                    try {
                        $null = Invoke-ConfluenceRequest -Method DELETE -Resource pages -ApiVersion 2 -Id $PageId -Query @{ purge = 'true' } -MaxQueryPages 1 -ErrorAction Stop
                        Write-Verbose "Purged page $PageId."
                    }
                    catch {
                        $detail = if ($trashError) { " Moving it to the trash also failed: $trashError" } else { '' }
                        Write-Error "Failed to purge page with ID $PageId. Error: $_$detail"
                    }
                }
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
