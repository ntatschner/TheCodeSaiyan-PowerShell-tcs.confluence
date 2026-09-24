function Clear-ConfluenceContext {
    <#
    .SYNOPSIS
        Removes the Confluence site and credential stored by Set-ConfluenceContext.

    .DESCRIPTION
        Clear-ConfluenceContext forgets the connection and the in-memory credential for the current
        session, and clears the cached space key to space ID lookups. REST commands report that no
        context is set until Set-ConfluenceContext is run again. Nothing is written to disk.

    .EXAMPLE
        Clear-ConfluenceContext

        Signs the session out of Confluence.

    .OUTPUTS
        None.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param ()

    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        $target = if ($script:ConfluenceContext) { $script:ConfluenceContext.ConnectionBaseURL } else { 'Confluence context' }
        if (-not $PSCmdlet.ShouldProcess($target, 'Clear Confluence context')) {
            return
        }
        $script:ConfluenceContext = $null
        $script:ConfluenceCredential = $null
        $script:ConfluenceSpaceIdCache = $null
        Write-Verbose 'Confluence context cleared.'
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        throw
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
