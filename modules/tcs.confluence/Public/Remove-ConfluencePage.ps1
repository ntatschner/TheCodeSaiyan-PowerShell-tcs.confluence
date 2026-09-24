function Remove-ConfluencePage {
    <#
    .SYNOPSIS
        Deletes a Confluence page.

    .DESCRIPTION
        Remove-ConfluencePage deletes the page with the given ID through the Confluence v2 pages API.
        The page moves to the space trash. Because this changes the site, you are asked to confirm
        unless you pass -Confirm:$false; -WhatIf shows what would be deleted.

    .PARAMETER PageId
        The ID of the page to delete. Accepts pipeline input by property name (id).

    .EXAMPLE
        Remove-ConfluencePage -PageId 123456

        Asks for confirmation, then deletes page 123456.

    .EXAMPLE
        (Get-ConfluencePage -SpaceKey DOCS -Search 'Draft*').Results | Remove-ConfluencePage -WhatIf

        Shows which draft pages would be deleted without deleting them.

    .OUTPUTS
        None.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, HelpMessage = 'The ID of the page to remove.')]
        [Alias('id')]
        [string]$PageId
    )

    process {
        if ($PSCmdlet.ShouldProcess("Confluence page $PageId", 'Remove')) {
            try {
                $null = Invoke-ConfluenceRequest -Method DELETE -Resource pages -Id $PageId -MaxQueryPages 1 -ErrorAction Stop
                Write-Verbose "Removed page with ID $PageId."
            }
            catch {
                Write-Error "Failed to remove page with ID $PageId. Error: $_"
            }
        }
    }
}
