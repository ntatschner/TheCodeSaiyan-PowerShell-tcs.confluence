function ConvertTo-ConfluenceHTML {
    <#
    .SYNOPSIS
        Converts simple Markdown to HTML for a Confluence page body.

    .DESCRIPTION
        ConvertTo-ConfluenceHTML performs a lightweight, regex-based Markdown conversion. Supported:
        headings (# to ######), **bold**, *italic*, "- " bullet lists, fenced code blocks (```), which
        may span several lines and whose content is HTML-encoded and left untouched by the other
        rules, and table rows (| a | b |), each converted to a single-cell table row.

        It is not a full Markdown parser; nested lists, links, images and inline code are left as they
        are.

    .PARAMETER InputContent
        The Markdown text to convert.

    .PARAMETER InputFormat
        The format of the input. Only Markdown is supported.

    .EXAMPLE
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# Title`n- one`n- two"

        Returns <h1>Title</h1> followed by <ul><li>one</li><li>two</li></ul>.

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

    switch ($InputFormat) {
        'Markdown' {
            $HtmlContent = $InputContent -replace "`r`n", "`n"

            # Fenced code blocks first (they may span lines); keep them out of the other rules
            $codeBlocks = New-Object -TypeName System.Collections.Generic.List[string]
            $codeMatches = [regex]::Matches($HtmlContent, '(?s)```[^\n`]*\n?(.*?)```')
            for ($index = $codeMatches.Count - 1; $index -ge 0; $index--) {
                $match = $codeMatches[$index]
                $code = [System.Net.WebUtility]::HtmlEncode($match.Groups[1].Value.TrimEnd("`n"))
                $codeBlocks.Insert(0, "<pre><code>$code</code></pre>")
                $HtmlContent = $HtmlContent.Remove($match.Index, $match.Length).Insert($match.Index, "@@CODEBLOCK$index@@")
            }

            # Headings
            $HtmlContent = $HtmlContent -replace '(?m)^###### (.*)$', '<h6>$1</h6>'
            $HtmlContent = $HtmlContent -replace '(?m)^##### (.*)$', '<h5>$1</h5>'
            $HtmlContent = $HtmlContent -replace '(?m)^#### (.*)$', '<h4>$1</h4>'
            $HtmlContent = $HtmlContent -replace '(?m)^### (.*)$', '<h3>$1</h3>'
            $HtmlContent = $HtmlContent -replace '(?m)^## (.*)$', '<h2>$1</h2>'
            $HtmlContent = $HtmlContent -replace '(?m)^# (.*)$', '<h1>$1</h1>'

            # Bold, then italic
            $HtmlContent = $HtmlContent -replace '\*\*(.+?)\*\*', '<strong>$1</strong>'
            $HtmlContent = $HtmlContent -replace '\*(.+?)\*', '<em>$1</em>'

            # Unordered lists (consecutive items share one list)
            $HtmlContent = $HtmlContent -replace '(?m)^- (.*)$', '<ul><li>$1</li></ul>'
            $HtmlContent = $HtmlContent -replace '</ul>\n<ul>', ''

            # Table rows
            $HtmlContent = $HtmlContent -replace '(?m)^\|(.*)\|$', '<table><tr><td>$1</td></tr></table>'

            for ($index = 0; $index -lt $codeBlocks.Count; $index++) {
                $HtmlContent = $HtmlContent.Replace("@@CODEBLOCK$index@@", $codeBlocks[$index])
            }
            return $HtmlContent
        }
    }
}
