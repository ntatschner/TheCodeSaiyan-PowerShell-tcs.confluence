function ConvertTo-ConfluenceHTML {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "The input content to be converted.")]
        [string]$InputContent,

        [Parameter(Mandatory = $true, HelpMessage = "The format of the input content. Currently supported: 'Markdown'.")]
        [ValidateSet("Markdown")]
        [string]$InputFormat
    )

    process {
        switch ($InputFormat) {
            "Markdown" {
                try {
                    # Convert headings
                    $HtmlContent = $InputContent -replace '(^|\n)###### (.*)', '$1<h6>$2</h6>'
                    $HtmlContent = $HtmlContent -replace '(^|\n)##### (.*)', '$1<h5>$2</h5>'
                    $HtmlContent = $HtmlContent -replace '(^|\n)#### (.*)', '$1<h4>$2</h4>'
                    $HtmlContent = $HtmlContent -replace '(^|\n)### (.*)', '$1<h3>$2</h3>'
                    $HtmlContent = $HtmlContent -replace '(^|\n)## (.*)', '$1<h2>$2</h2>'
                    $HtmlContent = $HtmlContent -replace '(^|\n)# (.*)', '$1<h1>$2</h1>'

                    # Convert bold text
                    $HtmlContent = $HtmlContent -replace '\*\*(.*?)\*\*', '<strong>$1</strong>'

                    # Convert italic text
                    $HtmlContent = $HtmlContent -replace '\*(.*?)\*', '<em>$1</em>'

                    # Convert unordered lists
                    $HtmlContent = $HtmlContent -replace '(^|\n)- (.*)', '$1<ul><li>$2</li></ul>'
                    $HtmlContent = $HtmlContent -replace '</ul>\n<ul>', ''

                    # Convert code blocks
                    $HtmlContent = $HtmlContent -replace '```(.*?)```', '<pre><code>$1</code></pre>'
                    # Convert tables
                    $HtmlContent = $HtmlContent -replace '\|(.*)\|', '<table><tr><td>$1</td></tr></table>'
                    return $HtmlContent
                } catch {
                    Write-Error "Failed to convert Markdown to HTML. Error: $_"
                }
            }
            default {
                Write-Error "Unsupported input format: $InputFormat"
            }
        }
    }
}
