function Join-ConfluenceContent {
    <#
    .SYNOPSIS
        Joins blocks of Confluence content with a separator.

    .DESCRIPTION
        Join-ConfluenceContent concatenates storage-format content blocks, such as the output of the
        New-ConfluenceContent* functions, with a horizontal rule, line break, space or tab between them.

    .PARAMETER ContentBlocks
        The content blocks to join, in order.

    .PARAMETER Separator
        The separator between blocks: NewLine (<br />, default), HorizontalRule (<hr />),
        Space (&nbsp;) or Tab (&emsp;).

    .EXAMPLE
        Join-ConfluenceContent -ContentBlocks (New-ConfluenceContentHeader -Header 'Intro' -Level 1), '<p>Text</p>' -Separator HorizontalRule

        Returns <h1>Intro</h1><hr /><p>Text</p>.

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The blocks of content to join together')]
        [ValidateNotNullOrEmpty()]
        [string[]]$ContentBlocks,

        [Parameter(HelpMessage = 'The separator to use between content blocks')]
        [ValidateSet('HorizontalRule', 'NewLine', 'Space', 'Tab')]
        [string]$Separator = 'NewLine'
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
        $separatorText = switch ($Separator) {
            'HorizontalRule' { '<hr />' }
            'NewLine' { '<br />' }
            'Space' { '&nbsp;' }
            'Tab' { '&emsp;' }
        }

        return ($ContentBlocks -join $separatorText)
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
