function Remove-ConfluencePageLabel {
    <#
    .SYNOPSIS
        Removes labels from a Confluence page.

    .DESCRIPTION
        Remove-ConfluencePageLabel removes one or more labels from a page through the Confluence v1
        API (DELETE /wiki/rest/api/content/<id>/label?name=<label>); the v2 API has no endpoint to
        remove labels. Supports -WhatIf and -Confirm.

    .PARAMETER PageId
        The ID of the page. Accepts pipeline input by property name (id).

    .PARAMETER Label
        The label names to remove.

    .EXAMPLE
        Remove-ConfluencePageLabel -PageId 123456 -Label draft

        Removes the draft label from page 123456.

    .OUTPUTS
        None.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, HelpMessage = 'The ID of the page.')]
        [Alias('id')]
        [string]$PageId,

        [Parameter(Mandatory, HelpMessage = 'The labels to remove.')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Label
    )

    begin {
        $telemetry = Start-TcsTelemetry
    }

    process {
        try {
            foreach ($name in $Label) {
                if (-not $PSCmdlet.ShouldProcess("Confluence page $PageId", "Remove label '$name'")) {
                    continue
                }
                try {
                    $null = Invoke-ConfluenceRequest -Method DELETE -URIPath "/wiki/rest/api/content/$PageId/label" -Query @{ name = $name } -ErrorAction Stop
                    Write-Verbose "Removed label '$name' from page $PageId."
                }
                catch {
                    Write-Error "Failed to remove label '$name' from page $PageId. Error: $_"
                }
            }
        }
        catch {
            Complete-TcsTelemetry -Token $telemetry -ErrorRecord $_
            throw
        }
    }

    end {
        Complete-TcsTelemetry -Token $telemetry
    }
}
