function Get-ConfluencePageLabel {
    <#
    .SYNOPSIS
        Gets the labels of a Confluence page.

    .DESCRIPTION
        Get-ConfluencePageLabel returns the labels of a page through the Confluence v2 API
        (GET /wiki/api/v2/pages/<id>/labels), reading every result page. Each label has id, name and
        prefix (global, my or team).

    .PARAMETER PageId
        The ID of the page. Accepts pipeline input by property name (id), for example from
        Get-ConfluencePage.

    .PARAMETER Prefix
        Only return labels with this prefix: global, my or team.

    .EXAMPLE
        Get-ConfluencePageLabel -PageId 123456 | Select-Object -ExpandProperty name

        Lists the label names of page 123456.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. One object per label.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, HelpMessage = 'The ID of the page.')]
        [Alias('id')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Only return labels with this prefix.')]
        [ValidateSet('global', 'my', 'team')]
        [string]$Prefix
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
            $query = @{ limit = 250 }
            if ($Prefix) { $query.prefix = $Prefix }
            try {
                $response = Invoke-ConfluenceRequest -Method GET -URIPath "/wiki/api/v2/pages/$PageId/labels" -Query $query -All -ErrorAction Stop
                foreach ($label in @($response.Results)) { $label }
            }
            catch {
                Write-Error "Failed to get the labels of page $PageId. Error: $_"
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
