function New-ConfluenceContentExpand {
    <#
    .SYNOPSIS
        Creates a Confluence expand macro (a collapsible section).

    .DESCRIPTION
        New-ConfluenceContentExpand returns the storage-format markup for the Confluence "expand"
        macro: a section that is collapsed until the reader clicks its title. The title is escaped.
        The content is escaped and wrapped in a paragraph; use -Raw to insert storage-format markup,
        for example a table or code block built with the other New-ConfluenceContent* functions.

    .PARAMETER Title
        The text of the clickable title. Confluence shows "Click here to expand..." when it is empty.

    .PARAMETER Content
        The body of the section: plain text, or storage-format markup with -Raw.

    .PARAMETER Raw
        Insert -Content as storage-format markup without escaping or wrapping it in a paragraph.

    .EXAMPLE
        New-ConfluenceContentExpand -Title 'Details' -Content (New-ConfluenceContentTable -TableData $rows) -Raw

        Returns a collapsed "Details" section that contains a table.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = 'The title of the expand section.')]
        [string]$Title,

        [Parameter(Mandatory, HelpMessage = 'The content of the expand section.')]
        [AllowEmptyString()]
        [string]$Content,

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
        $bodyMarkup = if ($Raw) { $Content } else { '<p>' + (ConvertTo-ConfluenceXmlText -Text $Content) + '</p>' }
        $markup = '<ac:structured-macro ac:name="expand">'
        if ($Title) {
            $markup += '<ac:parameter ac:name="title">' + (ConvertTo-ConfluenceXmlText -Text $Title) + '</ac:parameter>'
        }
        $markup += "<ac:rich-text-body>$bodyMarkup</ac:rich-text-body></ac:structured-macro>"
        return $markup
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
