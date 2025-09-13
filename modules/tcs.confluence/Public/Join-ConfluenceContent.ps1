function Join-ConfluenceContent {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = "The blocks of content to join together")]
        [ValidateNotNullOrWhiteSpace()]    
        [string[]]$ContentBlocks,

        [Parameter(HelpMessage = "The separator to use between content blocks")]
        [ValidateNotNullOrWhiteSpace()]
        [ValidateSet("HorizontalRule", "NewLine", "Space", "Tab")]
        [string]$Separator = "NewLine"
    )

    $separatorText = switch ($Separator) {
        "HorizontalRule" { '<hr />' }
        "NewLine" { '<br />' }
        "Space" { '&nbsp;' }
        "Tab" { '&emsp;' }
    }

    $JoinedContent = $ContentBlocks -join $separatorText
    return $JoinedContent
}
        
