function New-ConfluenceContentLink {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = "The text to use instead of the full URL.", ParameterSetName = "Url")]
        [string]$LinkText,

        [Parameter(Mandatory, HelpMessage = "The URL to link to.", ParameterSetName = "Url")]
        [string]$Url,

        [Parameter(Mandatory, HelpMessage = "Format any urls in the text block.", ParameterSetName = "TextBlock")]
        [string]$TextBlock
    )

    begin {
        $LinkText = if ($LinkText) { $LinkText } else { $Url }
    }
    process {
        if ([string]::IsNullOrEmpty($TextBlock) -eq $false) {
            $Pattern = '\b((http|https):\/\/)?((www\.)?([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})(\/[a-zA-Z0-9-._~:\/?#[\]@!$&''()*+,;=]*)?\b'
            $LinkHtml = $TextBlock -replace $Pattern, { New-ConfluenceContentLink -Url $_.Value }
        }
        else {
            $LinkHtml = "<a href='$Url'>$LinkText</a>"
        }
        return $LinkHtml
    }
}
