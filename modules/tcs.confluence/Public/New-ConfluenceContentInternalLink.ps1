function New-ConfluenceContentInternalLink {
    <#
    .SYNOPSIS
        Creates a link to a Confluence page, optionally to a heading on that page.

    .DESCRIPTION
        New-ConfluenceContentInternalLink returns an <a> element that links to a Confluence page given
        by its URL or by its page ID. With -PageId the page URL is looked up with Get-ConfluencePage,
        which needs a Confluence context (Set-ConfluenceContext).

        With -HeadingLink the link points to a heading on the page, using the Confluence anchor format
        "#<PageTitle>-<Heading>" with spaces removed. The page title comes from -PageTitle or, when not
        given, is looked up from the page ID in the URL.

    .PARAMETER InternalLinkURL
        The full URL of the page to link to.

    .PARAMETER PageTitle
        The title of the page, used to build the heading anchor.

    .PARAMETER PageId
        The ID of the page to link to.

    .PARAMETER HeadingLink
        The heading on the target page to link to.

    .PARAMETER LinkText
        The text shown for the link. Defaults to the URL.

    .EXAMPLE
        New-ConfluenceContentInternalLink -InternalLinkURL 'https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook' -PageTitle 'Runbook' -HeadingLink 'Restart steps' -LinkText 'Restart steps'

        Returns a link to the "Restart steps" heading of the Runbook page.

    .EXAMPLE
        New-ConfluenceContentInternalLink -PageId 123456 -LinkText 'See the runbook'

        Looks up page 123456 and returns a link to it.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding(DefaultParameterSetName = 'InternalLinkURL')]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = 'URL to the page to link to.', ParameterSetName = 'InternalLinkURL')]
        [string]$InternalLinkURL,

        [Parameter(HelpMessage = 'Title of the page to link to.', ParameterSetName = 'InternalLinkURL')]
        [string]$PageTitle,

        [Parameter(Mandatory, HelpMessage = 'Page ID to link to.', ParameterSetName = 'InternalLinkPageId')]
        [string]$PageId,

        [Parameter(HelpMessage = 'The hash text to link on the internal page.')]
        [string]$HeadingLink,

        [Parameter(HelpMessage = 'The text to use instead of the full URL.')]
        [string]$LinkText
    )

    if ($PSCmdlet.ParameterSetName -eq 'InternalLinkURL') {
        $Link = $InternalLinkURL.TrimEnd('/') + '/'
    }
    else {
        $context = Get-ConfluenceContext
        if (-not $context) {
            Write-Error 'No Confluence context is set. Run Set-ConfluenceContext first.'
            return
        }
        $page = (Get-ConfluencePage -PageId $PageId -ErrorAction Stop).Results | Select-Object -First 1
        $webUi = if ($page -and $page._links) { $page._links.webui } else { $null }
        if (-not $webUi) {
            Write-Error "Could not find the web URL of page $PageId."
            return
        }
        $Link = $context.ConnectionBaseURL + '/wiki' + $webUi
        if (-not $PageTitle) { $PageTitle = $page.title }
    }
    Write-Verbose "Link: $Link"

    if ($HeadingLink) {
        if (-not $PageTitle) {
            $parsedId = [regex]::Match($Link, '/pages/(\d+)(/|$)').Groups[1].Value
            Write-Verbose "Parsed page ID: $parsedId"
            if (-not $parsedId) {
                Write-Error "Could not find a page ID in '$Link'. Pass -PageTitle to link to a heading."
                return
            }
            $PageTitle = ((Get-ConfluencePage -PageId $parsedId -ErrorAction Stop).Results | Select-Object -First 1).title
        }
        $Link = $Link + '#' + ($PageTitle -replace ' ', '') + '-' + ($HeadingLink -replace ' ', '')
        Write-Verbose "Link: $Link"
    }

    if ($LinkText) {
        return (New-ConfluenceContentLink -Url $Link -LinkText $LinkText)
    }
    return (New-ConfluenceContentLink -Url $Link)
}
