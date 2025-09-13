function Set-ConfluenceContext {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "The base URL of the Confluence instance.")]
        [ValidatePattern('^https://')]
        [string]$ConfluenceUrl,
        [Parameter(Mandatory)][Alias("EmailAddress")][string]$Username,
        [Parameter(Mandatory)][Alias("PAT")][string]$PersonalAccessToken,
        [ValidateSet("v1","v2")][string]$ApiVersion = "v2"
    )

    # --- Normalize base URL ---
    $raw = $ConfluenceUrl.Trim().TrimEnd('/')
    $normalized = ($raw -replace '(?i)/wiki/?(api/v2|rest/api)?$', '')
    $apiSuffix = if ($ApiVersion -eq 'v1') { '/wiki/rest/api' } else { '/wiki/api/v2' }
    $ConnectionURI = "$normalized$apiSuffix"

    # FIX (braces)
    $basicPair  = "${Username}:${PersonalAccessToken}"
    $authHeader = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($basicPair))

    $ContextParams = @{
        OriginalConnectionURL = $raw
        ConnectionBaseURL     = $normalized
        ConnectionURI         = $ConnectionURI
        Username              = $Username
        PersonalAccessToken   = ('*' * 8)
        AuthorizationHeader   = @{
            Authorization = $authHeader
            ContentType   = "application/json"
            Accept        = "application/json"
        }
    }
    $global:ConfluenceContext = [pscustomobject]$ContextParams
    Write-Verbose "Confluence context set. Base='$normalized' API='$apiSuffix'"
}
