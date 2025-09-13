function New-ConfluenceContentInfo {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = "The title of the info block.")]
        [string]$Title,

        [Parameter(Mandatory, HelpMessage = "The content of the info block.")]
        [string]$Content,

        [Parameter(HelpMessage = "The type of info block to create.")]
        [ValidateSet("info", "tip", "note", "warning", "error")]
        [string]$Type = "info"
    )

    begin {
        $InfoBlock = @{
            info    = "info"
            tip     = "tip"
            note    = "note"
            warning = "warning"
            error   = "error"
        }
    }
    process {
        $InfoBlockHtml = @"
<div class="aui-message aui-message-$($InfoBlock[$Type])">
    <p class="title">
        <span class="aui-icon icon-$($InfoBlock[$Type])"></span>
        $Title
    </p>
    <p>$Content</p>
</div>
"@
        return $InfoBlockHtml
    }
}
