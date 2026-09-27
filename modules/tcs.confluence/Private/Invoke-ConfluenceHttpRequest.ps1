function Invoke-ConfluenceHttpRequest {
    <#
    .SYNOPSIS
        Sends one HTTP request to Confluence and returns the status code, body and Retry-After value.

    .DESCRIPTION
        Wraps Invoke-WebRequest so HTTP error responses are returned instead of thrown on both
        Windows PowerShell 5.1 and PowerShell 7. Returns an object with StatusCode, StatusDescription,
        Content and RetryAfter. Transport failures (DNS, TLS, timeouts) are rethrown and not retried.
        The request headers, which carry the credential, are never written to any output stream.

        A 429 Too Many Requests response is retried up to -MaxRetry times with tcs.core
        Invoke-WithRetry. The wait is the Retry-After header (seconds or an HTTP date, capped at 60
        seconds) or, without that header, 2, 4, 8 ... seconds. No other status is retried. When the
        last retry is still throttled, the 429 response is returned like any other error response.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE')]
        [string]$Method,

        [string]$Body,

        [byte[]]$BodyBytes,

        [string]$ContentType = 'application/json; charset=utf-8',

        [hashtable]$AdditionalHeaders,

        [ValidateRange(0, 10)]
        [int]$MaxRetry = 4
    )

    # Shared with the script blocks below: the number of requests sent, the 429 response of the
    # latest request and the wait before the next retry
    $state = @{ Attempt = 0; Throttled = $null; DelaySeconds = 0 }

    $sendRequest = {
        $state.Attempt++
        $state.Throttled = $null
        $response = Invoke-ConfluenceHttpRequestOnce -Uri $Uri -Method $Method -Body $Body -BodyBytes $BodyBytes -ContentType $ContentType -AdditionalHeaders $AdditionalHeaders
        if ($response.StatusCode -ne 429) {
            return $response
        }

        # Throw the 429 so that Invoke-WithRetry retries it. The wait is passed on as the error's
        # Retry-After value, so Invoke-WithRetry sleeps exactly that long.
        $state.Throttled = $response
        $state.DelaySeconds = Get-ConfluenceRetryDelay -RetryAfter $response.RetryAfter -Attempt $state.Attempt
        $exception = New-Object -TypeName System.Exception -ArgumentList "Confluence returned $($response.StatusCode) $($response.StatusDescription)."
        $throttledResponse = [pscustomobject]@{
            StatusCode        = $response.StatusCode
            StatusDescription = $response.StatusDescription
            Headers           = @{ 'Retry-After' = [string]$state.DelaySeconds }
        }
        Add-Member -InputObject $exception -MemberType NoteProperty -Name Response -Value $throttledResponse
        throw $exception
    }

    # Only a 429 is retried; errors without an HTTP response (transport failures) are not
    $isThrottled = {
        param ($ErrorRecord)
        $detail = Get-HttpErrorDetail -ErrorRecord $ErrorRecord
        return ($null -ne $state.Throttled -and $null -ne $detail -and $detail.StatusCode -eq 429)
    }

    $onRetry = {
        param ($RetryException, $RetryAttempt)
        Write-Verbose "Confluence returned 429 Too Many Requests; retry $RetryAttempt of $MaxRetry in $($state.DelaySeconds) second(s)."
    }

    $retryParams = @{
        ScriptBlock       = $sendRequest
        MaxRetries        = $MaxRetry
        DelaySeconds      = 0
        MaxDelaySeconds   = 60
        RetryOnStatusCode = 429
        ShouldRetry       = $isThrottled
        OnRetry           = $onRetry
    }
    try {
        Invoke-WithRetry @retryParams
    }
    catch {
        if ($null -ne $state.Throttled) {
            return $state.Throttled
        }
        throw
    }
}

function Get-ConfluenceRetryDelay {
    <#
    .SYNOPSIS
        Returns the number of seconds to wait before retrying a 429 response.

    .DESCRIPTION
        Uses the Retry-After value (a number of seconds or an HTTP date) when there is one, otherwise
        2 to the power of the attempt number (2, 4, 8 ... seconds). The result is between 0 and 60.
    #>
    [CmdletBinding()]
    [OutputType([double])]
    param (
        [string]$RetryAfter,

        [Parameter(Mandatory)]
        [int]$Attempt
    )

    $delaySeconds = [math]::Pow(2, $Attempt)
    if ($RetryAfter) {
        $seconds = 0
        $date = [datetime]::MinValue
        if ([int]::TryParse($RetryAfter, [ref]$seconds)) {
            $delaySeconds = $seconds
        }
        elseif ([datetime]::TryParse($RetryAfter, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$date)) {
            $delaySeconds = [math]::Ceiling(($date - [datetime]::UtcNow).TotalSeconds)
        }
    }
    return [math]::Min([math]::Max($delaySeconds, 0), 60)
}

function Invoke-ConfluenceHttpRequestOnce {
    <#
    .SYNOPSIS
        Sends a single HTTP request without retrying. Used by Invoke-ConfluenceHttpRequest.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [string]$Method,

        [string]$Body,

        [byte[]]$BodyBytes,

        [string]$ContentType,

        [hashtable]$AdditionalHeaders
    )

    if ($null -eq $script:ConfluenceContext -or $null -eq $script:ConfluenceCredential) {
        throw 'No Confluence context is set. Run Set-ConfluenceContext first.'
    }
    # Built per request and never stored, so the token is not kept in plain text in session state
    $headers = New-BasicAuthHeader -Credential $script:ConfluenceCredential
    $headers.Accept = 'application/json'
    if ($AdditionalHeaders) {
        foreach ($key in @($AdditionalHeaders.get_Keys())) { $headers[$key] = $AdditionalHeaders[$key] }
    }
    $requestParams = @{
        Uri             = $Uri
        Method          = $Method
        Headers         = $headers
        ContentType     = $ContentType
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if ($null -ne $BodyBytes -and $BodyBytes.Length -gt 0) {
        $requestParams.Body = $BodyBytes
    }
    elseif (-not [string]::IsNullOrEmpty($Body)) {
        # Send UTF-8 bytes so non-ASCII page content survives on Windows PowerShell 5.1
        $requestParams.Body = [System.Text.Encoding]::UTF8.GetBytes($Body)
    }
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $requestParams.SkipHttpErrorCheck = $true
    }

    try {
        $response = Invoke-WebRequest @requestParams
        return [pscustomobject]@{
            StatusCode        = [int]$response.StatusCode
            StatusDescription = [string]$response.StatusDescription
            Content           = [string]$response.Content
            RetryAfter        = (Get-ConfluenceRetryAfterValue -Headers $response.Headers)
        }
    }
    catch {
        # Windows PowerShell 5.1 throws on HTTP errors; recover the status and body from the error.
        # An error without an HTTP response is a transport failure and is rethrown.
        $detail = Get-HttpErrorDetail -ErrorRecord $_
        if ($null -eq $detail) {
            throw
        }
        $responseHeaders = $null
        if ($null -ne $detail.Response -and $detail.Response.PSObject.Properties['Headers']) {
            $responseHeaders = $detail.Response.Headers
        }
        return [pscustomobject]@{
            StatusCode        = [int]$detail.StatusCode
            StatusDescription = [string]$detail.StatusDescription
            Content           = [string]$detail.Body
            RetryAfter        = (Get-ConfluenceRetryAfterValue -Headers $responseHeaders)
        }
    }
}

function Get-ConfluenceRetryAfterValue {
    <#
    .SYNOPSIS
        Returns the Retry-After header value from a response header collection, or nothing.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        $Headers
    )

    if ($null -eq $Headers) { return }
    # Dictionaries (PowerShell 7, and Windows PowerShell for successful responses) throw for a
    # missing key; WebHeaderCollection returns $null
    if ($Headers -is [System.Collections.IDictionary] -and -not ([System.Collections.IDictionary]$Headers).Contains('Retry-After')) { return }
    $value = $null
    try {
        $value = $Headers['Retry-After']
    }
    catch {
        Write-Verbose 'The response headers could not be read.'
    }
    if ($null -eq $value) { return }
    # PowerShell 7 returns a string[] per header; Windows PowerShell returns a string
    return [string](@($value)[0])
}
