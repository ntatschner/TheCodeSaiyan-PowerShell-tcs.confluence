function Get-ConfluenceAttachment {
    <#
    .SYNOPSIS
        Gets the attachments of a Confluence page.

    .DESCRIPTION
        Get-ConfluenceAttachment returns the attachments of a page through the Confluence v2 API
        (GET /wiki/api/v2/pages/<id>/attachments), reading every result page. Each attachment has
        id, title, mediaType, fileSize, version and downloadLink properties.

    .PARAMETER PageId
        The ID of the page. Accepts pipeline input by property name (id).

    .PARAMETER FileName
        Only return the attachment with this file name.

    .EXAMPLE
        Get-ConfluenceAttachment -PageId 123456 | Select-Object title, fileSize

        Lists the files attached to page 123456.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. One object per attachment.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, HelpMessage = 'The ID of the page.')]
        [Alias('id')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Only return the attachment with this file name.')]
        [string]$FileName
    )

    begin {
        $telemetry = Start-TcsTelemetry
        $lastError = $null
    }

    process {
        $completed = $false
        try {
            $query = @{ limit = 250 }
            if ($FileName) { $query.filename = $FileName }
            try {
                $response = Invoke-ConfluenceRequest -Method GET -URIPath "/wiki/api/v2/pages/$PageId/attachments" -Query $query -All -ErrorAction Stop
                foreach ($attachment in @($response.Results)) { $attachment }
            }
            catch {
                Write-Error "Failed to get the attachments of page $PageId. Error: $_"
            }
            $completed = $true
        }
        catch {
            $lastError = $_
            throw
        }
        finally {
            # A stopped pipeline (Select-Object -First) or a terminating error skips the end block
            if (-not $completed) {
                Complete-TcsTelemetry -Token $telemetry -ErrorRecord $lastError
            }
        }
    }

    end {
        Complete-TcsTelemetry -Token $telemetry -ErrorRecord $lastError
    }
}
