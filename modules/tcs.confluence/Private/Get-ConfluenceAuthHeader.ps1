function Get-ConfluenceAuthHeader {
    <#
    .SYNOPSIS
        Builds the HTTP headers for a Confluence request from the in-memory credential.

    .DESCRIPTION
        Returns a new hashtable with a Basic Authorization header built from the credential stored by
        Set-ConfluenceContext. The header is built per request and never stored, so the token is not
        kept in plain text in session state. Throws when no context has been set.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param ()

    if ($null -eq $script:ConfluenceContext -or $null -eq $script:ConfluenceCredential) {
        throw 'No Confluence context is set. Run Set-ConfluenceContext first.'
    }

    $networkCredential = $script:ConfluenceCredential.GetNetworkCredential()
    $pair = '{0}:{1}' -f $networkCredential.UserName, $networkCredential.Password
    $encoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($pair))

    return @{
        Authorization = "Basic $encoded"
        Accept        = 'application/json'
    }
}
