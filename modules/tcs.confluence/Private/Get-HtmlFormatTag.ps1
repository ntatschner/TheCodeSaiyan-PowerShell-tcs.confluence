function Get-HtmlFormatTag {
    <#
    .SYNOPSIS
        Returns the opening or closing HTML tags for a list of text formats.

    .DESCRIPTION
        Maps Bold, Italic, Underline and Strikethrough to strong, em, u and s. With -Close the closing
        tags are returned in reverse order so the tags nest correctly.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [AllowNull()]
        [AllowEmptyCollection()]
        [string[]]$Format,

        [switch]$Close
    )

    $tagNames = @{
        Bold          = 'strong'
        Italic        = 'em'
        Underline     = 'u'
        Strikethrough = 's'
    }
    $tags = @(foreach ($item in @($Format)) {
            if ($item -and $tagNames.ContainsKey($item)) { $tagNames[$item] }
        })
    if ($Close) {
        [array]::Reverse($tags)
        return (($tags | ForEach-Object { "</$_>" }) -join '')
    }
    return (($tags | ForEach-Object { "<$_>" }) -join '')
}
