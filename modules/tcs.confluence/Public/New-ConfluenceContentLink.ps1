function New-ConfluenceContentLink {
    <#
    .SYNOPSIS
        Creates a hyperlink, or turns every URL in a block of text into a hyperlink.

    .DESCRIPTION
        With -Url, New-ConfluenceContentLink returns an <a> element that shows -LinkText (or the URL
        when no text is given). With -TextBlock, every web address found in the text is replaced by a
        link to itself.

    .PARAMETER LinkText
        The text shown for the link. Defaults to the URL.

    .PARAMETER Url
        The address to link to.

    .PARAMETER TextBlock
        Text in which every URL is converted to a link.

    .EXAMPLE
        New-ConfluenceContentLink -Url 'https://example.com' -LinkText 'Example'

        Returns <a href='https://example.com'>Example</a>.

    .EXAMPLE
        New-ConfluenceContentLink -TextBlock 'See https://example.com/docs for details.'

        Returns the sentence with the address converted to a link.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding(DefaultParameterSetName = 'Url')]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = 'The text to use instead of the full URL.', ParameterSetName = 'Url')]
        [string]$LinkText,

        [Parameter(Mandatory, HelpMessage = 'The URL to link to.', ParameterSetName = 'Url')]
        [string]$Url,

        [Parameter(Mandatory, HelpMessage = 'Format any urls in the text block.', ParameterSetName = 'TextBlock')]
        [string]$TextBlock
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
        if ($PSCmdlet.ParameterSetName -eq 'TextBlock') {
            $Pattern = '\b((http|https):\/\/)?((www\.)?([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})(\/[a-zA-Z0-9-._~:\/?#[\]@!$&''()*+,;=]*)?\b'
            # A MatchEvaluator works on Windows PowerShell 5.1, where -replace does not accept a script block
            $evaluator = [System.Text.RegularExpressions.MatchEvaluator] {
                param($Match)
                "<a href='$($Match.Value)'>$($Match.Value)</a>"
            }
            return [regex]::Replace($TextBlock, $Pattern, $evaluator)
        }

        if (-not $LinkText) { $LinkText = $Url }
        return "<a href='$Url'>$LinkText</a>"
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
