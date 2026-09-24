function ConvertTo-ConfluenceHTML {
    <#
    .SYNOPSIS
        Converts simple Markdown to HTML for a Confluence page body.

    .DESCRIPTION
        ConvertTo-ConfluenceHTML performs a lightweight, line-based Markdown conversion. Supported:
        headings (# to ######), **bold**, *italic*, bullet lists ("- " or "* "; consecutive items share
        one list), fenced code blocks (```), whose content is kept exactly and left untouched by the
        other rules, and tables: consecutive | a | b | lines become one table, a |---|---| separator
        row is skipped and marks the row above it as the header row.

        All text is escaped (& < >), so the result is well-formed storage format and HTML in the
        Markdown is shown as text. Other lines are passed through as escaped text, one per line.

        It is not a full Markdown parser; nested lists, links, images and inline code are left as
        text.

    .PARAMETER InputContent
        The Markdown text to convert.

    .PARAMETER InputFormat
        The format of the input. Only Markdown is supported.

    .EXAMPLE
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# Title`n- one`n- two"

        Returns <h1>Title</h1> followed by <ul><li>one</li><li>two</li></ul>.

    .EXAMPLE
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "| Name | Count |`n|---|---|`n| a | 1 |"

        Returns one table with a header row (Name, Count) and one data row.

    .OUTPUTS
        System.String
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'The input content to be converted.')]
        [AllowEmptyString()]
        [string]$InputContent,

        [Parameter(Mandatory = $true, HelpMessage = "The format of the input content. Currently supported: 'Markdown'.")]
        [ValidateSet('Markdown')]
        [string]$InputFormat
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
        switch ($InputFormat) {
            'Markdown' {
                $inline = {
                    param([string]$Text)
                    $html = ConvertTo-ConfluenceXmlText -Text $Text
                    $html = $html -replace '\*\*(.+?)\*\*', '<strong>$1</strong>'
                    $html = $html -replace '\*(.+?)\*', '<em>$1</em>'
                    return $html
                }
                $splitRow = {
                    param([string]$Line)
                    $cells = $Line.Trim()
                    if ($cells.StartsWith('|')) { $cells = $cells.Substring(1) }
                    if ($cells.EndsWith('|')) { $cells = $cells.Substring(0, $cells.Length - 1) }
                    # "\|" is a literal pipe inside a cell
                    return , @($cells.Replace('\|', [string][char]0) -split '\|' | ForEach-Object { $_.Replace([string][char]0, '|').Trim() })
                }
                $separatorPattern = '^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$'

                $lines = ($InputContent -replace "`r`n", "`n") -split "`n"
                $blocks = New-Object -TypeName System.Collections.Generic.List[string]
                $index = 0
                while ($index -lt $lines.Count) {
                    $line = $lines[$index]
                    if ($line -match '^\s*```') {
                        # Fenced code block: everything up to the closing fence is kept as-is
                        $code = New-Object -TypeName System.Collections.Generic.List[string]
                        $index++
                        while ($index -lt $lines.Count -and $lines[$index] -notmatch '^\s*```') {
                            $code.Add($lines[$index])
                            $index++
                        }
                        $index++
                        $blocks.Add('<pre><code>' + (ConvertTo-ConfluenceXmlText -Text ($code -join "`n")) + '</code></pre>')
                        continue
                    }
                    if ($line -match '^(#{1,6})\s+(.*)$') {
                        $level = $Matches[1].Length
                        $blocks.Add("<h$level>" + (& $inline $Matches[2]) + "</h$level>")
                        $index++
                        continue
                    }
                    if ($line -match '^\s*[-*]\s+') {
                        $items = ''
                        while ($index -lt $lines.Count -and $lines[$index] -match '^\s*[-*]\s+(.*)$') {
                            $items += '<li>' + (& $inline $Matches[1]) + '</li>'
                            $index++
                        }
                        $blocks.Add("<ul>$items</ul>")
                        continue
                    }
                    if ($line -match '^\s*\|.*\|\s*$') {
                        $tableLines = New-Object -TypeName System.Collections.Generic.List[string]
                        while ($index -lt $lines.Count -and $lines[$index] -match '^\s*\|.*\|\s*$') {
                            $tableLines.Add($lines[$index])
                            $index++
                        }
                        $hasHeader = $tableLines.Count -gt 1 -and $tableLines[1] -match $separatorPattern
                        $rows = ''
                        for ($rowIndex = 0; $rowIndex -lt $tableLines.Count; $rowIndex++) {
                            if ($tableLines[$rowIndex] -match $separatorPattern) { continue }
                            $cellTag = if ($hasHeader -and $rowIndex -eq 0) { 'th' } else { 'td' }
                            $cells = foreach ($cell in (& $splitRow $tableLines[$rowIndex])) { "<$cellTag>" + (& $inline $cell) + "</$cellTag>" }
                            $rows += '<tr>' + ($cells -join '') + '</tr>'
                        }
                        $blocks.Add("<table><tbody>$rows</tbody></table>")
                        continue
                    }
                    $blocks.Add((& $inline $line))
                    $index++
                }
                $HtmlContent = $blocks -join "`n"
                return $HtmlContent
            }
        }
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
