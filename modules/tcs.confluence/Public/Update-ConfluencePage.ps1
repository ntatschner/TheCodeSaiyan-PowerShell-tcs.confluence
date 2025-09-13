function Update-ConfluencePage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, HelpMessage = "The page ID of the page to update.")]
        [string]$PageId,

        [Parameter(HelpMessage = "Space key / id (optional if unchanged).")]
        [string]$SpaceKey,

        [Parameter(Mandatory, HelpMessage = "Title of the page.")]
        [string]$Title,

        [Parameter(HelpMessage = "Status: draft or current.")]
        [ValidateSet('draft','current')]
        [string]$Status = 'current',

        [Parameter(Mandatory, HelpMessage = "Storage format content.")]
        [string]$Content,

        [Parameter(Mandatory, HelpMessage = "New version number (increment previous).")]
        [int]$Version,

        [string]$VersionMessage = "Programmatically Updated"
    )

    process {
        $body = @{
            id      = $PageId
            title   = $Title
            status  = $Status
            body    = @{
                value          = $Content
                representation = 'storage'
            }
            version = @{
                number  = $Version
                message = $VersionMessage
            }
        }
        if ($SpaceKey) { $body.spaceId = $SpaceKey }

        try {
            $resp = Invoke-ConfluenceRequest -Method PUT -Resource pages -Id $PageId -Body ($body | ConvertTo-Json -Depth 10) -ErrorAction Stop
            # Invoke-ConfluenceRequest returns object with .Results (array) or raw payload
            if ($resp.Results) { return $resp.Results[0] }
            return $resp
        } catch {
            Write-Error "Failed to update page '$PageId'. Error: $_"
        }
    }
}
