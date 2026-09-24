function Resolve-ConfluenceSpaceId {
    <#
    .SYNOPSIS
        Returns the numeric space ID for a space key, using a per-session cache.

    .DESCRIPTION
        A numeric value is returned unchanged. A space key is looked up once with the v2 spaces API
        (GET /wiki/api/v2/spaces?keys=<key>) and cached for the session, so repeated requests for the
        same space do not look it up again. Set-ConfluenceContext and Clear-ConfluenceContext clear
        the cache. Throws when the key cannot be resolved.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string]$Space
    )

    $Space = $Space.Trim()
    if ($Space -match '^\d+$') { return $Space }
    if ($null -eq $script:ConfluenceSpaceIdCache) {
        $script:ConfluenceSpaceIdCache = New-Object -TypeName 'System.Collections.Generic.Dictionary[string,string]' -ArgumentList ([System.StringComparer]::Ordinal)
    }
    if ($script:ConfluenceSpaceIdCache.ContainsKey($Space)) {
        Write-Verbose "Space key '$Space' resolved to space ID '$($script:ConfluenceSpaceIdCache[$Space])' (cached)."
        return $script:ConfluenceSpaceIdCache[$Space]
    }
    if ($null -eq $script:ConfluenceContext -or $null -eq $script:ConfluenceCredential) {
        throw 'No Confluence context is set. Run Set-ConfluenceContext first.'
    }

    $uri = '{0}/wiki/api/v2/spaces?keys={1}' -f $script:ConfluenceContext.ConnectionBaseURL, [System.Uri]::EscapeDataString($Space)
    $response = Invoke-ConfluenceHttpRequest -Uri $uri -Method GET
    if ($response.StatusCode -lt 200 -or $response.StatusCode -gt 299) {
        throw "Could not resolve space key '$Space' to a space ID. Status: $($response.StatusCode) $($response.StatusDescription)"
    }
    $payload = $response.Content | ConvertFrom-Json
    $spaceId = @($payload.results | Where-Object { $_.key -eq $Space } | ForEach-Object { "$($_.id)" }) | Select-Object -First 1
    if (-not $spaceId) {
        throw "Space key '$Space' was not found."
    }
    $script:ConfluenceSpaceIdCache[$Space] = $spaceId
    Write-Verbose "Space key '$Space' resolved to space ID '$spaceId'."
    return $spaceId
}
