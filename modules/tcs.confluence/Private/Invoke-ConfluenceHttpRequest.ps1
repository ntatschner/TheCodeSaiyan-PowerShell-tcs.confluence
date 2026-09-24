function Invoke-ConfluenceHttpRequest {
    <#
    .SYNOPSIS
        Sends one HTTP request to Confluence and returns the status code and body.

    .DESCRIPTION
        Wraps Invoke-WebRequest so HTTP error responses are returned instead of thrown on both
        Windows PowerShell 5.1 and PowerShell 7. Returns an object with StatusCode, StatusDescription
        and Content. Transport failures (DNS, TLS, timeouts) are rethrown. The request headers, which
        carry the credential, are never written to any output stream.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE')]
        [string]$Method,

        [string]$Body
    )

    $requestParams = @{
        Uri             = $Uri
        Method          = $Method
        Headers         = (Get-ConfluenceAuthHeader)
        ContentType     = 'application/json; charset=utf-8'
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if (-not [string]::IsNullOrEmpty($Body)) {
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
        }
    }
}
