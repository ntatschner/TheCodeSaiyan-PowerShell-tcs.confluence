function New-ConfluenceContentInternalLink {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = "URL to the page to link to.", ParameterSetName = "InternalLinkURL")]
        $InternalLinkURL,

        [Parameter(HelpMessage = "Title of the page to link to.", ParameterSetName = "InternalLinkURL")]
        $PageTitle,

        [Parameter(Mandatory, HelpMessage = "Page ID to link to.", ParameterSetName = "InternalLinkPageId")]
        $PageId,

        [Parameter(HelpMessage = "The hash text to link on the internal page.")]
        $HeadingLink,

        [Parameter(HelpMessage = "The text to use instead of the full URL.")]
        $LinkText
    )
    process {
        if ($PSCmdlet.ParameterSetName -eq "InternalLinkURL") {
            $Link = $($InternalLinkURL.TrimEnd("/")) + "/"
            Write-Verbose "Link: $Link"
        }
        else {
            $Link = $($ConfluenceContext.ConnectionBaseURL) + "/wiki" + $($(Get-ConfluencePage -PageId $PageId).Results._links.webui)
            Write-Verbose "Link: $Link"
        }
        if ($HeadingLink) {
            $regex = [regex]::new(".*\/(\d*)\/.*")
            $parseLink = $regex.Match($Link).Groups[1].Value
            Write-Verbose "Parsed LinkId: $parseLink"
            if ($PageTitle) {
                $Link = $Link + "#" + $($PageTitle -replace " ", "") + "-" + $HeadingLink.replace(" ", "")
            }
            else {
                $Page = Get-ConfluencePage -PageId $parseLink
                $PageTitleLink = "#" + $($Page.Results.title -replace " ", "") + "-" + $HeadingLink.replace(" ", "")
            }
            $Link = $Link + $PageTitleLink
            Write-Verbose "Link: $Link"
        }
        if ($LinkText) {
            return $(New-ConfluenceContentLink -URL $Link -LinkText $LinkText)
        }
        else {
            return $(New-ConfluenceContentLink -URL $Link)
        }
    }
}
