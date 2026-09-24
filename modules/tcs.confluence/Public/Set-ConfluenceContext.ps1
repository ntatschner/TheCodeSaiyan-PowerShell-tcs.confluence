function Set-ConfluenceContext {
    <#
    .SYNOPSIS
        Sets the Confluence site and credential used by the other tcs.confluence commands.

    .DESCRIPTION
        Set-ConfluenceContext stores the Confluence base URL and your credential for the current
        PowerShell session. The URL is normalised (a trailing /wiki, /wiki/api/v2 or /wiki/rest/api is
        removed) and the API endpoint for the chosen version is derived from it.

        The credential is kept in memory only, as a PSCredential (the token is held in a SecureString).
        It is never written to disk, never returned by Get-ConfluenceContext and never written to the
        verbose, warning or error streams. The Authorization header is built for each request.

        Use Get-ConfluenceContext to see the current connection details and Clear-ConfluenceContext to
        forget them. Setting a context also clears the cached space key to space ID lookups.

    .PARAMETER ConfluenceUrl
        The base URL of the Confluence site, for example https://contoso.atlassian.net. Only https URLs
        are accepted so the credential is never sent in clear text.

    .PARAMETER Username
        The account e-mail address (Confluence Cloud) or user name used with the API token.

    .PARAMETER PersonalAccessToken
        The API token or personal access token as plain text. Kept for backward compatibility; prefer
        -Credential so the token is not visible in your command history.

    .PARAMETER Credential
        A PSCredential whose user name is the account e-mail address and whose password is the API
        token, for example from Get-Credential or a secret store.

    .PARAMETER ApiVersion
        The default REST API version: v2 (default) or v1. It sets the derived ConnectionURI and the
        version Invoke-ConfluenceRequest -Resource uses when the call does not pass -ApiVersion. The
        page, space, label and attachment commands always use the API version they are written for.

    .EXAMPLE
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Credential (Get-Credential)

        Prompts for the e-mail address and API token and stores them for the session.

    .EXAMPLE
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net/wiki' -Username 'me@contoso.com' -PersonalAccessToken $token

        Uses a token held in a variable. The URL is normalised to https://contoso.atlassian.net.

    .OUTPUTS
        None.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low', DefaultParameterSetName = 'Token')]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'The base URL of the Confluence instance.')]
        [ValidatePattern('^https://')]
        [string]$ConfluenceUrl,

        [Parameter(Mandatory = $true, ParameterSetName = 'Token')]
        [Alias('EmailAddress')]
        [ValidateNotNullOrEmpty()]
        [string]$Username,

        [Parameter(Mandatory = $true, ParameterSetName = 'Token')]
        [Alias('PAT')]
        [ValidateNotNullOrEmpty()]
        [string]$PersonalAccessToken,

        [Parameter(Mandatory = $true, ParameterSetName = 'Credential')]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential,

        [ValidateSet('v1', 'v2')]
        [string]$ApiVersion = 'v2'
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
        # --- Normalise the base URL ---
        $raw = $ConfluenceUrl.Trim().TrimEnd('/')
        $normalized = ($raw -replace '(?i)/wiki/?(api/v2|rest/api)?$', '')
        $apiSuffix = if ($ApiVersion -eq 'v1') { '/wiki/rest/api' } else { '/wiki/api/v2' }
        $connectionUri = "$normalized$apiSuffix"

        if ($PSCmdlet.ParameterSetName -eq 'Token') {
            $secureToken = New-Object -TypeName System.Security.SecureString
            foreach ($character in $PersonalAccessToken.ToCharArray()) {
                $secureToken.AppendChar($character)
            }
            $secureToken.MakeReadOnly()
            $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $secureToken
        }

        if (-not $PSCmdlet.ShouldProcess($normalized, 'Set Confluence context')) {
            return
        }

        $script:ConfluenceCredential = $Credential
        $script:ConfluenceSpaceIdCache = $null
        $script:ConfluenceContext = [pscustomobject]@{
            OriginalConnectionURL = $raw
            ConnectionBaseURL     = $normalized
            ConnectionURI         = $connectionUri
            ApiVersion            = $ApiVersion
            Username              = $Credential.UserName
        }
        Write-Verbose "Confluence context set. Base='$normalized' API='$apiSuffix' User='$($Credential.UserName)'"
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
