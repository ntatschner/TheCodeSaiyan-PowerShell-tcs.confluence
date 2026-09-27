function Add-ConfluencePageLabel {
    <#
    .SYNOPSIS
        Adds labels to a Confluence page.

    .DESCRIPTION
        Add-ConfluencePageLabel adds one or more global labels to a page through the Confluence v1 API
        (POST /wiki/rest/api/content/<id>/label); the v2 API has no endpoint to add labels. Labels
        that the page already has are left as they are. Confluence stores labels in lower case and
        they cannot contain spaces. Returns the labels of the page after the change.

        Supports -WhatIf and -Confirm.

    .PARAMETER PageId
        The ID of the page. Accepts pipeline input by property name (id).

    .PARAMETER Label
        The label names to add.

    .EXAMPLE
        Add-ConfluencePageLabel -PageId 123456 -Label runbook, ops

        Adds the labels runbook and ops to page 123456.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. The labels of the page.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, HelpMessage = 'The ID of the page.')]
        [Alias('id')]
        [string]$PageId,

        [Parameter(Mandatory, HelpMessage = 'The labels to add.')]
        [ValidatePattern('^\S+$')]
        [string[]]$Label
    )

    begin {
        $telemetry = Start-TcsTelemetry
    }

    process {
        try {
            if (-not $PSCmdlet.ShouldProcess("Confluence page $PageId", "Add label(s) $($Label -join ', ')")) {
                return
            }
            $labels = @(foreach ($name in $Label) { [ordered]@{ prefix = 'global'; name = $name } })
            $body = ConvertTo-Json -InputObject $labels -Depth 3 -Compress
            try {
                $response = Invoke-ConfluenceRequest -Method POST -URIPath "/wiki/rest/api/content/$PageId/label" -Body $body -ErrorAction Stop
                foreach ($item in @($response.Results)) { $item }
            }
            catch {
                Write-Error "Failed to add labels to page $PageId. Error: $_"
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
