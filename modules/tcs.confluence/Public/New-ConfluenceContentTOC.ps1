function New-ConfluenceContentTOC {
    <#
    .SYNOPSIS
        Creates a Confluence table of contents macro.

    .DESCRIPTION
        New-ConfluenceContentTOC returns the storage-format markup for the Confluence "toc" macro,
        built from the headings on the page.

    .PARAMETER HorizontalList
        Show the table of contents as a horizontal (flat) list instead of a vertical list.

    .PARAMETER BulletPointStyle
        The list style: None (default), Mixed (Confluence default bullets), Bullet (disc), Circle,
        Square or Number (decimal).

    .PARAMETER HeadersFromLevel
        The lowest heading level to include, 1 to 6. Default 1.

    .PARAMETER HeadersToLevel
        The highest heading level to include, 1 to 6. Default 6.

    .PARAMETER IncludeSectionNumbers
        Number the entries as an outline (1, 1.1, 1.2 ...).

    .EXAMPLE
        New-ConfluenceContentTOC -HeadersFromLevel 2 -HeadersToLevel 3 -IncludeSectionNumbers

        Returns a numbered table of contents of the h2 and h3 headings.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = 'Format the table of contents as a horizontal list.')]
        [switch]$HorizontalList,

        [Parameter(HelpMessage = "Bullet point style for the table of contents.`n Default: None")]
        [ValidateSet('None', 'Mixed', 'Bullet', 'Circle', 'Square', 'Number')]
        [string]$BulletPointStyle = 'None',

        [Parameter(HelpMessage = "The level of headers to include from in the table of contents.`n Default: 1")]
        [ValidateRange(1, 6)]
        [int]$HeadersFromLevel = 1,

        [Parameter(HelpMessage = "The level of headers to include to in the table of contents.`n Default: 6")]
        [ValidateRange(1, 6)]
        [int]$HeadersToLevel = 6,

        [Parameter(HelpMessage = 'Whether to include section numbers in the table of contents.')]
        [switch]$IncludeSectionNumbers
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
        if ($HeadersFromLevel -gt $HeadersToLevel) {
            Write-Error "HeadersFromLevel ($HeadersFromLevel) must not be greater than HeadersToLevel ($HeadersToLevel)."
            return
        }

        # Confluence toc macro: type = list | flat, outline = section numbering, style = CSS list-style-type
        $TypeValue = if ($HorizontalList) { 'flat' } else { 'list' }
        $OutlineValue = if ($IncludeSectionNumbers) { 'true' } else { 'false' }
        $StyleMap = @{
            None   = 'none'
            Bullet = 'disc'
            Circle = 'circle'
            Square = 'square'
            Number = 'decimal'
        }

        $TOCMacro = "<ac:structured-macro ac:name='toc'>"
        $TOCMacro += "<ac:parameter ac:name='maxLevel'>$HeadersToLevel</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='minLevel'>$HeadersFromLevel</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='type'>$TypeValue</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='outline'>$OutlineValue</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='printable'>true</ac:parameter>"
        if ($StyleMap.ContainsKey($BulletPointStyle)) {
            $TOCMacro += "<ac:parameter ac:name='style'>$($StyleMap[$BulletPointStyle])</ac:parameter>"
        }
        $TOCMacro += '</ac:structured-macro>'
        return $TOCMacro
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
