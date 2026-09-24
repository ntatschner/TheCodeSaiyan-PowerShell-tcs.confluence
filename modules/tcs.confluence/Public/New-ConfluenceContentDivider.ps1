function New-ConfluenceContentDivider {
    <#
    .SYNOPSIS
        Creates a divider (horizontal rule or blank paragraph) for a Confluence page.

    .DESCRIPTION
        New-ConfluenceContentDivider returns a self-closed <hr/> element (storage format is XHTML),
        optionally with a style class, or an empty paragraph for vertical space.

    .PARAMETER Type
        The divider type: line (default), space, default, dashed, dotted, double or gradient.

    .EXAMPLE
        New-ConfluenceContentDivider -Type space

        Returns <p>&#160;</p> (a non-breaking space).

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = 'The type of divider to create.')]
        [ValidateSet('line', 'space', 'default', 'dashed', 'dotted', 'double', 'gradient')]
        [string]$Type = 'line'
    )

    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        $Divider = @{
            line     = '<hr/>'
            space    = '<p>&#160;</p>'
            default  = "<hr class='default'/>"
            dashed   = "<hr class='dashed'/>"
            dotted   = "<hr class='dotted'/>"
            double   = "<hr class='double'/>"
            gradient = "<hr class='gradient'/>"
        }
        return $Divider[$Type]
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
