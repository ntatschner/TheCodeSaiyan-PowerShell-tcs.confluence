function Get-ConfluenceContext {
    <#
    .SYNOPSIS
        Returns the Confluence connection set by Set-ConfluenceContext, without the credential.

    .DESCRIPTION
        Get-ConfluenceContext returns the base URL, API endpoint, API version and user name stored by
        Set-ConfluenceContext for the current session. The token is never returned. Returns nothing
        when no context has been set.

        This replaces reading the $global:ConfluenceContext variable used by tcs.confluence 0.0.x,
        which exposed the Authorization header to every script in the session.

    .EXAMPLE
        Get-ConfluenceContext

        Shows the current connection, for example ConnectionBaseURL https://contoso.atlassian.net.

    .EXAMPLE
        if (-not (Get-ConfluenceContext)) { Set-ConfluenceContext -ConfluenceUrl $url -Credential (Get-Credential) }

        Sets a context only when none exists yet.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param ()

    if ($null -eq $script:ConfluenceContext) {
        Write-Verbose 'No Confluence context is set. Run Set-ConfluenceContext first.'
        return
    }

    return [pscustomobject]@{
        OriginalConnectionURL = $script:ConfluenceContext.OriginalConnectionURL
        ConnectionBaseURL     = $script:ConfluenceContext.ConnectionBaseURL
        ConnectionURI         = $script:ConfluenceContext.ConnectionURI
        ApiVersion            = $script:ConfluenceContext.ApiVersion
        Username              = $script:ConfluenceContext.Username
        HasCredential         = ($null -ne $script:ConfluenceCredential)
    }
}
