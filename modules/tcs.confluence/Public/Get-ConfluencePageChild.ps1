function Get-ConfluencePageChild {
    <#
    .SYNOPSIS
        Gets the child pages of a Confluence page, or with -Recurse all pages below it.

    .DESCRIPTION
        Get-ConfluencePageChild returns the child pages of a page through the Confluence v2 API
        (GET /wiki/api/v2/pages/<id>/children), reading every result page. With -Recurse the children
        of every child are read as well, level by level, so every descendant page is returned (parents
        before their children). Each page gets a parentId property if the API did not return one.

    .PARAMETER PageId
        The ID of the parent page. Accepts pipeline input by property name (id).

    .PARAMETER Recurse
        Return all descendant pages, not only the direct children.

    .EXAMPLE
        Get-ConfluencePageChild -PageId 123456 | Select-Object id, title

        Lists the direct child pages of page 123456.

    .EXAMPLE
        Get-ConfluencePageChild -PageId 123456 -Recurse | Measure-Object

        Counts every page below page 123456.

    .OUTPUTS
        System.Management.Automation.PSCustomObject. One object per page.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName, HelpMessage = 'The ID of the parent page.')]
        [Alias('id')]
        [string]$PageId,

        [Parameter(HelpMessage = 'Return all descendant pages.')]
        [switch]$Recurse
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
            $queue = New-Object -TypeName System.Collections.Generic.Queue[string]
            $seen = New-Object -TypeName System.Collections.Generic.HashSet[string]
            $queue.Enqueue($PageId)
            $null = $seen.Add($PageId)
            while ($queue.Count -gt 0) {
                $parentId = $queue.Dequeue()
                try {
                    $response = Invoke-ConfluenceRequest -Method GET -URIPath "/wiki/api/v2/pages/$parentId/children" -Query @{ limit = 250 } -All -ErrorAction Stop
                }
                catch {
                    Write-Error "Failed to get the child pages of page $parentId. Error: $_"
                    continue
                }
                foreach ($child in @($response.Results)) {
                    if (@($child.PSObject.Properties.Name) -notcontains 'parentId') {
                        $child | Add-Member -NotePropertyName parentId -NotePropertyValue $parentId
                    }
                    $child
                    if ($Recurse -and $child.id -and $seen.Add([string]$child.id)) {
                        $queue.Enqueue([string]$child.id)
                    }
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
