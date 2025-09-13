function Remove-ConfluencePage {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "The ID of the page to remove.")]
        [string]$PageId
    )

    begin {
        if (-not (Get-Variable -Scope Global -Name "ConfluenceContext" -ErrorAction SilentlyContinue)) {
            Write-Error "ConfluenceContext variable not found. Run Set-ConfluenceContext first."
            return
        }
    }

    process {
        if ($pscmdlet.ShouldProcess("page with ID $PageId", "Remove")) {
            try {
                $resource = "pages/$PageId"
                Invoke-ConfluenceRequest -Resource $resource -Method 'DELETE' -ErrorAction Stop
                Write-Verbose "Successfully submitted request to remove page with ID $PageId."
            }
            catch {
                Write-Error "Failed to remove page with ID $PageId. Error: $_"
            }
        }
    }
}
