function New-ConfluenceContentInfo {
    <#
    .SYNOPSIS
        Creates an info, tip, note, warning or error message block.

    .DESCRIPTION
        New-ConfluenceContentInfo returns an AUI message block (a div with the aui-message classes)
        with a title and content. Title and content are inserted as given, so they may contain markup.

    .PARAMETER Title
        The title of the block.

    .PARAMETER Content
        The body of the block.

    .PARAMETER Type
        The block type: info (default), tip, note, warning or error.

    .EXAMPLE
        New-ConfluenceContentInfo -Title 'Heads up' -Content 'Maintenance on Friday.' -Type warning

        Returns a warning message block.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The title of the info block.')]
        [string]$Title,

        [Parameter(Mandatory, HelpMessage = 'The content of the info block.')]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(HelpMessage = 'The type of info block to create.')]
        [ValidateSet('info', 'tip', 'note', 'warning', 'error')]
        [string]$Type = 'info'
    )

    $typeName = $Type.ToLowerInvariant()
    $InfoBlockHtml = @"
<div class="aui-message aui-message-$typeName">
    <p class="title">
        <span class="aui-icon icon-$typeName"></span>
        $Title
    </p>
    <p>$Content</p>
</div>
"@
    return $InfoBlockHtml
}
