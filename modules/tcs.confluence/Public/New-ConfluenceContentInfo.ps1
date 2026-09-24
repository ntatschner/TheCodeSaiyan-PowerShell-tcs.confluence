function New-ConfluenceContentInfo {
    <#
    .SYNOPSIS
        Creates a Confluence info, tip, note or warning panel macro.

    .DESCRIPTION
        New-ConfluenceContentInfo returns the storage-format markup for the Confluence info, tip, note
        or warning macro (<ac:structured-macro ac:name="info"> with a title parameter and a
        rich-text body). Confluence has no "error" panel macro, so -Type error produces the warning
        macro.

        The title is always escaped. The content is escaped and wrapped in a paragraph; use -Raw to
        insert storage-format markup (for example the output of other New-ConfluenceContent*
        functions) as the body instead.

        Before 0.2.0 this function returned an AUI message div, which Confluence does not render as a
        panel.

    .PARAMETER Title
        The title of the panel. An empty title shows no title.

    .PARAMETER Content
        The body of the panel: plain text, or storage-format markup with -Raw.

    .PARAMETER Type
        The panel type: info (default), tip, note, warning or error (rendered as warning).

    .PARAMETER Raw
        Insert -Content as storage-format markup without escaping or wrapping it in a paragraph.

    .EXAMPLE
        New-ConfluenceContentInfo -Title 'Heads up' -Content 'Maintenance on Friday.' -Type warning

        Returns a warning panel.

    .EXAMPLE
        New-ConfluenceContentInfo -Title 'Steps' -Content (New-ConfluenceContentCodeBlock -Content 'Restart-Service web') -Raw

        Returns an info panel whose body is a code block.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The title of the info block.')]
        [AllowEmptyString()]
        [string]$Title,

        [Parameter(Mandatory, HelpMessage = 'The content of the info block.')]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(HelpMessage = 'The type of info block to create.')]
        [ValidateSet('info', 'tip', 'note', 'warning', 'error')]
        [string]$Type = 'info',

        [Parameter(HelpMessage = 'Insert the content as storage-format markup without escaping it.')]
        [switch]$Raw
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
        # Confluence storage format has info, tip, note and warning panels; "error" maps to warning
        $macroName = $Type.ToLowerInvariant()
        if ($macroName -eq 'error') { $macroName = 'warning' }
        $bodyMarkup = if ($Raw) { $Content } else { '<p>' + (ConvertTo-ConfluenceXmlText -Text $Content) + '</p>' }

        $InfoBlockHtml = "<ac:structured-macro ac:name=`"$macroName`">"
        if ($Title) {
            $InfoBlockHtml += '<ac:parameter ac:name="title">' + (ConvertTo-ConfluenceXmlText -Text $Title) + '</ac:parameter>'
        }
        $InfoBlockHtml += "<ac:rich-text-body>$bodyMarkup</ac:rich-text-body></ac:structured-macro>"
        return $InfoBlockHtml
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
