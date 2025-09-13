function New-ConfluenceContentCodeBlock {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "The code block content.")]
        [string]$Content,

        [Parameter(HelpMessage = "The language of the code block.")]
        [string]$Language = "none",

        [Parameter(HelpMessage = "The theme of the code block.")]
        [ValidateSet("Default", "Midnight", "Eclipse", "Emacs")]
        [string]$Theme = "Default",

        [Parameter(HelpMessage = "Whether to show line numbers.")]
        [switch]$LineNumbers,

        [Parameter(HelpMessage = "Whether to collapse the code block.")]
        [bool]$Collapse = $false
    )

    process {
        $LineNumbersValue = if ($LineNumbers) { "true" } else { "false" }
        $CollapseValue = if ($Collapse) { "true" } else { "false" }

        $CodeBlockHtml = "<ac:structured-macro ac:name='code'>"
        $CodeBlockHtml += "<ac:parameter ac:name='theme'>$Theme</ac:parameter>"
        $CodeBlockHtml += "<ac:parameter ac:name='linenumbers'>$LineNumbersValue</ac:parameter>"
        $CodeBlockHtml += "<ac:parameter ac:name='collapse'>$CollapseValue</ac:parameter>"
        $CodeBlockHtml += "<ac:parameter ac:name='language'>$Language</ac:parameter>"
        $CodeBlockHtml += "<ac:plain-text-body><![CDATA[$Content]]></ac:plain-text-body>"
        $CodeBlockHtml += "</ac:structured-macro>"

        return $CodeBlockHtml
    }
}
