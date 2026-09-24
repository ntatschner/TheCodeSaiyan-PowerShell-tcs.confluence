function New-ConfluenceContentInternalLink {
    <#
    .SYNOPSIS
        Creates a link to a Confluence page, optionally to a heading on that page.

    .DESCRIPTION
        By default New-ConfluenceContentInternalLink returns a Confluence storage-format page link:

            <ac:link><ri:page ri:content-title="Title" ri:space-key="KEY"/><ac:link-body>Text</ac:link-body></ac:link>

        Confluence resolves this link by page title, keeps it working when the page is moved or
        renamed, and shows it as a page link. Give the page with -PageTitle (and -SpaceKey for a page
        in another space), or with -PageId, in which case the title and space key are looked up with
        Get-ConfluencePage (this needs Set-ConfluenceContext). -HeadingLink adds ac:anchor so the link
        points to a heading on the page.

        URL mode (the behaviour before 0.2.0): with -InternalLinkURL, or with -PageId and -AsUrl, an
        <a href> element pointing to the page URL is returned instead. With -HeadingLink the URL gets
        the Confluence anchor "#<PageTitle>-<Heading>" with spaces removed; the page title comes from
        -PageTitle or is looked up from the page ID in the URL.

        All values are escaped, so the result is always well-formed storage format.

    .PARAMETER PageTitle
        The title of the page to link to. With -InternalLinkURL it is only used to build the heading
        anchor.

    .PARAMETER SpaceKey
        The key of the space that contains the page. Leave it out to link to a page in the same space
        as the page that contains the link.

    .PARAMETER InternalLinkURL
        The full URL of the page to link to (URL mode).

    .PARAMETER PageId
        The ID of the page to link to. Its title and space key (or, with -AsUrl, its URL) are looked
        up.

    .PARAMETER AsUrl
        With -PageId, return an <a href> link to the page URL instead of a storage-format page link.

    .PARAMETER HeadingLink
        The heading on the target page to link to.

    .PARAMETER LinkText
        The text shown for the link. Defaults to the page title (URL mode: the URL).

    .EXAMPLE
        New-ConfluenceContentInternalLink -PageTitle 'Runbook' -SpaceKey OPS -HeadingLink 'Restart steps' -LinkText 'Restart steps'

        Returns <ac:link ac:anchor="Restart steps"><ri:page ri:content-title="Runbook" ri:space-key="OPS"/><ac:link-body>Restart steps</ac:link-body></ac:link>.

    .EXAMPLE
        New-ConfluenceContentInternalLink -PageId 123456 -LinkText 'See the runbook'

        Looks up page 123456 and returns a storage-format link to it.

    .EXAMPLE
        New-ConfluenceContentInternalLink -InternalLinkURL 'https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook' -PageTitle 'Runbook' -HeadingLink 'Restart steps'

        Returns an <a href> link to the "Restart steps" heading of the Runbook page (URL mode).

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding(DefaultParameterSetName = 'PageTitle')]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = 'Title of the page to link to.', ParameterSetName = 'PageTitle')]
        [Parameter(HelpMessage = 'Title of the page to link to.', ParameterSetName = 'InternalLinkURL')]
        [string]$PageTitle,

        [Parameter(HelpMessage = 'Key of the space that contains the page.', ParameterSetName = 'PageTitle')]
        [string]$SpaceKey,

        [Parameter(Mandatory, HelpMessage = 'URL to the page to link to.', ParameterSetName = 'InternalLinkURL')]
        [string]$InternalLinkURL,

        [Parameter(Mandatory, HelpMessage = 'Page ID to link to.', ParameterSetName = 'InternalLinkPageId')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Return an <a href> link to the page URL.', ParameterSetName = 'InternalLinkPageId')]
        [switch]$AsUrl,

        [Parameter(HelpMessage = 'The heading on the page to link to.')]
        [string]$HeadingLink,

        [Parameter(HelpMessage = 'The text to show for the link.')]
        [string]$LinkText
    )

    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        $page = $null
        if ($PSCmdlet.ParameterSetName -eq 'InternalLinkPageId') {
            $context = Get-ConfluenceContext
            if (-not $context) {
                Write-Error 'No Confluence context is set. Run Set-ConfluenceContext first.'
                return
            }
            $page = Get-ConfluencePage -PageId $PageId -ErrorAction Stop | Select-Object -First 1
            if (-not $page) {
                Write-Error "Could not find page $PageId."
                return
            }
            $PageTitle = $page.title
        }

        if ($PSCmdlet.ParameterSetName -eq 'PageTitle' -or ($PSCmdlet.ParameterSetName -eq 'InternalLinkPageId' -and -not $AsUrl)) {
            if ($page) {
                # The space key is part of the page's web UI path: /spaces/<KEY>/pages/<id>/...
                $webUi = if ($page._links) { [string]$page._links.webui } else { '' }
                $keyMatch = [regex]::Match($webUi, '/spaces/([^/]+)/')
                if ($keyMatch.Success) { $SpaceKey = [System.Uri]::UnescapeDataString($keyMatch.Groups[1].Value) }
                if (-not $PageTitle) {
                    Write-Error "Could not find the title of page $PageId."
                    return
                }
            }
            $anchor = if ($HeadingLink) { ' ac:anchor="' + (ConvertTo-ConfluenceXmlText -Text $HeadingLink -Attribute) + '"' } else { '' }
            $spaceAttribute = if ($SpaceKey) { ' ri:space-key="' + (ConvertTo-ConfluenceXmlText -Text $SpaceKey -Attribute) + '"' } else { '' }
            $bodyText = if ($LinkText) { $LinkText } else { $PageTitle }
            return ('<ac:link{0}><ri:page ri:content-title="{1}"{2}/><ac:link-body>{3}</ac:link-body></ac:link>' -f
                $anchor, (ConvertTo-ConfluenceXmlText -Text $PageTitle -Attribute), $spaceAttribute, (ConvertTo-ConfluenceXmlText -Text $bodyText))
        }

        # URL mode
        if ($PSCmdlet.ParameterSetName -eq 'InternalLinkURL') {
            $Link = $InternalLinkURL.TrimEnd('/') + '/'
        }
        else {
            $webUi = if ($page._links) { $page._links.webui } else { $null }
            if (-not $webUi) {
                Write-Error "Could not find the web URL of page $PageId."
                return
            }
            $Link = $context.ConnectionBaseURL + '/wiki' + $webUi
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
                $PageTitle = (Get-ConfluencePage -PageId $parsedId -ErrorAction Stop | Select-Object -First 1).title
            }
            $Link = $Link + '#' + ($PageTitle -replace ' ', '') + '-' + ($HeadingLink -replace ' ', '')
            Write-Verbose "Link: $Link"
        }

        if (-not $LinkText) { $LinkText = $Link }
        return ("<a href='{0}'>{1}</a>" -f (ConvertTo-ConfluenceXmlText -Text $Link -Attribute), (ConvertTo-ConfluenceXmlText -Text $LinkText))
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        throw
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
