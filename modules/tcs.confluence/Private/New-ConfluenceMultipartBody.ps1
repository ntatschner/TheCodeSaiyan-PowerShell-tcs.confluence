function New-ConfluenceMultipartBody {
    <#
    .SYNOPSIS
        Builds a multipart/form-data request body for a Confluence attachment upload.

    .DESCRIPTION
        Returns an object with Bytes (the body) and ContentType (multipart/form-data with the
        boundary). The file is sent in the "file" part; -Comment and minorEdit are sent as text parts.
        Works on Windows PowerShell 5.1, which has no -Form parameter on Invoke-WebRequest.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a request body in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory)]
        [string]$FileName,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [byte[]]$FileBytes,

        [string]$Comment,

        [bool]$MinorEdit = $true
    )

    $boundary = '----tcsconfluence' + [guid]::NewGuid().ToString('N')
    $utf8 = New-Object -TypeName System.Text.UTF8Encoding -ArgumentList $false
    $stream = New-Object -TypeName System.IO.MemoryStream
    try {
        $writeText = {
            param([string]$Text)
            $bytes = $utf8.GetBytes($Text)
            $stream.Write($bytes, 0, $bytes.Length)
        }
        $safeName = $FileName.Replace('"', '%22').Replace("`r", '').Replace("`n", '')
        & $writeText ("--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"$safeName`"`r`nContent-Type: application/octet-stream`r`n`r`n")
        if ($FileBytes.Length -gt 0) { $stream.Write($FileBytes, 0, $FileBytes.Length) }
        & $writeText "`r`n"
        if ($Comment) {
            & $writeText ("--$boundary`r`nContent-Disposition: form-data; name=`"comment`"`r`nContent-Type: text/plain; charset=utf-8`r`n`r`n$Comment`r`n")
        }
        $minorEditValue = if ($MinorEdit) { 'true' } else { 'false' }
        & $writeText ("--$boundary`r`nContent-Disposition: form-data; name=`"minorEdit`"`r`n`r`n$minorEditValue`r`n")
        & $writeText "--$boundary--`r`n"
        return [pscustomobject]@{
            Bytes       = $stream.ToArray()
            ContentType = "multipart/form-data; boundary=$boundary"
        }
    }
    finally {
        $stream.Dispose()
    }
}
