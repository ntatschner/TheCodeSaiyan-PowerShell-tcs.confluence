function Invoke-ConfluenceHttpRequest {
    <#
    .SYNOPSIS
        Sends one HTTP request to Confluence and returns the status code, body and Retry-After value.

    .DESCRIPTION
        Wraps Invoke-WebRequest so HTTP error responses are returned instead of thrown on both
        Windows PowerShell 5.1 and PowerShell 7. Returns an object with StatusCode, StatusDescription,
        Content and RetryAfter. Transport failures (DNS, TLS, timeouts) are rethrown. The request
        headers, which carry the credential, are never written to any output stream.

        A 429 Too Many Requests response is retried up to -MaxRetry times. The wait is the
        Retry-After header (seconds or an HTTP date, capped at 60 seconds) or, without that header,
        2, 4, 8 ... seconds.
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

    $attempt = 0
    while ($true) {
        $response = Invoke-ConfluenceHttpRequestOnce -Uri $Uri -Method $Method -Body $Body -BodyBytes $BodyBytes -ContentType $ContentType -AdditionalHeaders $AdditionalHeaders
        if ($response.StatusCode -ne 429 -or $attempt -ge $MaxRetry) {
            return $response
        }
        $attempt++
        $delaySeconds = [math]::Pow(2, $attempt)
        if ($response.RetryAfter) {
            $seconds = 0
            $date = [datetime]::MinValue
            if ([int]::TryParse("$($response.RetryAfter)", [ref]$seconds)) {
                $delaySeconds = $seconds
            }
            elseif ([datetime]::TryParse("$($response.RetryAfter)", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$date)) {
                $delaySeconds = [math]::Ceiling(($date - [datetime]::UtcNow).TotalSeconds)
            }
        }
        $delaySeconds = [math]::Min([math]::Max($delaySeconds, 0), 60)
        Write-Verbose "Confluence returned 429 Too Many Requests; retry $attempt of $MaxRetry in $delaySeconds second(s)."
        Start-Sleep -Milliseconds ([int]($delaySeconds * 1000))
    }
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

    $headers = Get-ConfluenceAuthHeader
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
        $requestError = $_
        $webResponse = $null
        if ($requestError.Exception.PSObject.Properties.Name -contains 'Response') {
            $webResponse = $requestError.Exception.Response
        }
        if ($null -eq $webResponse) {
            throw
        }

        # Windows PowerShell 5.1 throws on HTTP errors; recover the status and body from the exception
        $content = ''
        if ($requestError.ErrorDetails -and $requestError.ErrorDetails.Message) {
            $content = $requestError.ErrorDetails.Message
        }
        else {
            try {
                $reader = New-Object -TypeName System.IO.StreamReader -ArgumentList $webResponse.GetResponseStream()
                $content = $reader.ReadToEnd()
                $reader.Dispose()
            }
            catch {
                Write-Verbose "Could not read the error response body: $($_.Exception.Message)"
            }
        }
        return [pscustomobject]@{
            StatusCode        = [int]$webResponse.StatusCode
            StatusDescription = [string]$webResponse.StatusDescription
            Content           = [string]$content
            RetryAfter        = (Get-ConfluenceRetryAfterValue -Headers $webResponse.Headers)
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
