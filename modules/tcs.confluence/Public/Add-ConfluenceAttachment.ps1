function Add-ConfluenceAttachment {
    <#
    .SYNOPSIS
        Uploads files as attachments to a Confluence page.

    .DESCRIPTION
        Add-ConfluenceAttachment uploads one or more files to a page through the Confluence v1 API
        (PUT /wiki/rest/api/content/<id>/child/attachment, multipart/form-data with the
        X-Atlassian-Token: no-check header); the v2 API has no upload endpoint. When the page already
        has an attachment with the same file name, a new version of that attachment is created.
        Returns the created or updated attachments.

        Supports -WhatIf and -Confirm.

    .PARAMETER PageId
        The ID of the page to attach the files to.

    .PARAMETER Path
        The files to upload. Accepts pipeline input (for example from Get-ChildItem).

    .PARAMETER Comment
        A comment stored with each attachment version.

    .PARAMETER NotifyWatchers
        Notify the page watchers. By default the upload is a minor edit and watchers are not notified.

    .EXAMPLE
        Add-ConfluenceAttachment -PageId 123456 -Path ./report.pdf -Comment 'Nightly report'

        Attaches report.pdf to page 123456, or adds a new version of it.

    .EXAMPLE
        Get-ChildItem ./out/*.png | Add-ConfluenceAttachment -PageId 123456

        Attaches every PNG file in ./out.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. One object per attachment.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The ID of the page.')]
        [string]$PageId,

        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'The files to upload.')]
        [Alias('FullName')]
        [string[]]$Path,

        [Parameter(HelpMessage = 'A comment for the attachment version.')]
        [string]$Comment,

        [Parameter(HelpMessage = 'Notify the page watchers.')]
        [switch]$NotifyWatchers
    )

    begin {
        $TelemetryArgs = @{
            ModuleName    = $MyInvocation.MyCommand.Module.Name
            ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
            CommandName   = $MyInvocation.MyCommand.Name
            ExecutionID   = [guid]::NewGuid().ToString()
        }
        Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
        $telemetryFailed = $false
    }

    process {
        try {
            if ($null -eq $script:ConfluenceContext -or $null -eq $script:ConfluenceCredential) {
                Write-Error 'No Confluence context is set. Run Set-ConfluenceContext first.'
                return
            }
            $uri = '{0}/wiki/rest/api/content/{1}/child/attachment' -f $script:ConfluenceContext.ConnectionBaseURL, [System.Uri]::EscapeDataString($PageId)
            foreach ($item in $Path) {
                $file = Get-Item -LiteralPath $item -ErrorAction SilentlyContinue
                if (-not $file -or $file.PSIsContainer) {
                    Write-Error "File not found: $item"
                    continue
                }
                if (-not $PSCmdlet.ShouldProcess("Confluence page $PageId", "Attach '$($file.Name)'")) {
                    continue
                }
                $multipart = New-ConfluenceMultipartBody -FileName $file.Name -FileBytes ([System.IO.File]::ReadAllBytes($file.FullName)) -Comment $Comment -MinorEdit (-not $NotifyWatchers)
                try {
                    $response = Invoke-ConfluenceHttpRequest -Uri $uri -Method PUT -BodyBytes $multipart.Bytes -ContentType $multipart.ContentType -AdditionalHeaders @{ 'X-Atlassian-Token' = 'no-check' } -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed to upload '$($file.Name)' to page $PageId. $($_.Exception.Message)"
                    continue
                }
                if ($response.StatusCode -lt 200 -or $response.StatusCode -gt 299) {
                    $errorText = $null
                    try {
                        $errorContent = $response.Content | ConvertFrom-Json -ErrorAction Stop
                        $errorText = (@($errorContent.errors | ForEach-Object { $_.title }) + @($errorContent.message) | Where-Object { $_ }) -join '; '
                    }
                    catch {
                        Write-Verbose 'The error response body is not JSON.'
                    }
                    Write-Error "Failed to upload '$($file.Name)' to page $PageId. Status: $($response.StatusCode) $($response.StatusDescription) Errors: $errorText"
                    continue
                }
                $payload = $response.Content | ConvertFrom-Json
                if (@($payload.PSObject.Properties.Name) -contains 'results') {
                    foreach ($attachment in @($payload.results)) { $attachment }
                }
                else {
                    $payload
                }
            }
        }
        catch {
            if (-not $telemetryFailed) {
                $telemetryFailed = $true
                Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
            }
            throw
        }
    }

    end {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
