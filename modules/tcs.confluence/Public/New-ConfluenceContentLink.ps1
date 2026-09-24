function New-ConfluenceContentLink {
    <#
    .SYNOPSIS
        Creates a hyperlink, or turns every URL in a block of text into a hyperlink.

    .DESCRIPTION
        With -Url, New-ConfluenceContentLink returns an <a> element that shows -LinkText (or the URL
        when no text is given). With -TextBlock, every address with a scheme (http://, https:// or
        mailto:) found in the text is replaced by a link to itself; file names such as report.pdf and
        bare e-mail addresses are left as text.

        The URL, the link text and the text block are escaped (& < > and quotes), so the result is
        always well-formed storage format.

    .PARAMETER LinkText
        The text shown for the link. Defaults to the URL.

    .PARAMETER Url
        The address to link to.

    .PARAMETER TextBlock
        Text in which every http, https or mailto address is converted to a link.

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
            return (ConvertTo-ConfluenceLinkedText -Text $TextBlock)
        }

        if (-not $LinkText) { $LinkText = $Url }
        return ("<a href='{0}'>{1}</a>" -f (ConvertTo-ConfluenceXmlText -Text $Url -Attribute), (ConvertTo-ConfluenceXmlText -Text $LinkText))
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
