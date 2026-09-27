function New-ConfluenceContentHeader {
    <#
    .SYNOPSIS
        Creates a heading (h1 to h6) for a Confluence page.

    .DESCRIPTION
        New-ConfluenceContentHeader returns a heading element with optional bold, italic, underline or
        strikethrough formatting. Formatting tags are nested correctly. The header text is escaped
        (& < > become entities) so it always produces valid storage format; use -Raw to insert
        markup that you have built yourself.

    .PARAMETER Header
        The heading text.

    .PARAMETER Level
        The heading level, 1 to 6.

    .PARAMETER StringFormatting
        Formatting to apply: Bold, Italic, Underline and/or Strikethrough.

    .PARAMETER Raw
        Insert -Header as given, without escaping. Only use it with trusted, well-formed markup.

    .EXAMPLE
        New-ConfluenceContentHeader -Header 'R&D' -Level 2

        Returns <h2>R&amp;D</h2>.

    .EXAMPLE
        New-ConfluenceContentHeader -Header 'Summary' -Level 2 -StringFormatting Bold, Italic

        Returns <h2><strong><em>Summary</em></strong></h2>.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'String to be used as the header.')]
        [string]$Header,

        [Parameter(Mandatory = $true, HelpMessage = 'The level of the header. Valid values: 1-6.')]
        [ValidateRange(1, 6)]
        [int]$Level,

        [Parameter(HelpMessage = 'The text formatting to be applied to the header.')]
        [ValidateSet('Bold', 'Italic', 'Underline', 'Strikethrough')]
        [string[]]$StringFormatting,

        [Parameter(HelpMessage = 'Insert the header text without escaping it.')]
        [switch]$Raw
    )

    $telemetry = Start-TcsTelemetry
    try {
        $open = Get-HtmlFormatTag -Format $StringFormatting
        $close = Get-HtmlFormatTag -Format $StringFormatting -Close
        $text = if ($Raw) { $Header } else { ConvertTo-ConfluenceXmlText -Text $Header }
        return "<h$Level>$open$text$close</h$Level>"
    }
    catch {
        Complete-TcsTelemetry -Token $telemetry -ErrorRecord $_
        throw
    }
    finally {
        Complete-TcsTelemetry -Token $telemetry
    }
}
