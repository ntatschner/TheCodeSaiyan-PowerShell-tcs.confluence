function New-ConfluenceContentTOC {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = "Format the table of contents as a horizontal list.")]
        [switch]$HorizontalList,

        [Parameter(HelpMessage = "Bullet point style for the table of contents.`n Default: None")]
        [ValidateSet("None", "Mixed", "Bullet", "Circle", "Square", "Number")]
        [string]$BulletPointStyle = "None",

        [Parameter(HelpMessage = "The level of headers to include from in the table of contents.`n Default: 1")]
        [ValidateRange(1, 6)]
        [string]$HeadersFromLevel = 1,

        [Parameter(HelpMessage = "The level of headers to include to in the table of contents.`n Default: 6")]
        [ValidateRange(1, 6)]
        [string]$HeadersToLevel = 6,

        [Parameter(HelpMessage = "Whether to include section numbers in the table of contents.")]
        [switch]$IncludeSectionNumbers
    )

    begin {
        $HorizontalListValue = if ($HorizontalList) { "true" } else { "false" }
        $IncludeSectionNumbersValue = if ($IncludeSectionNumbers) { "true" } else { "false" }
    }
    process {
        # Create Macro to embed the TOC
        $TOCMacro = "<ac:structured-macro ac:name='toc'>"
        $TOCMacro += "<ac:parameter ac:name='maxLevel'>$HeadersToLevel</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='minLevel'>$HeadersFromLevel</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='type'>list</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='outline'>$HorizontalListValue</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='include'>$IncludeSectionNumbersValue</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='printable'>true</ac:parameter>"
        $TOCMacro += "<ac:parameter ac:name='style'>$BulletPointStyle</ac:parameter>"
        $TOCMacro += "</ac:structured-macro>"
        return $TOCMacro
    }
}
