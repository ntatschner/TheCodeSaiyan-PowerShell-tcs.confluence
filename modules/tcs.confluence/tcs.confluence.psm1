#region load classes, then private and public functions
$ClassFiles = @(Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Classes') -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*.Tests.ps1' })
$Private = @(Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Private') -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*.Tests.ps1' })
$Public = @(Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Public') -Filter '*.ps1' -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*.Tests.ps1' })

foreach ($File in @($ClassFiles + $Private + $Public)) {
    try {
        . $File.FullName
    }
    catch {
        Write-Error -Message "Failed to import '$($File.FullName)': $_"
    }
}
#endregion

#region session state visible to the module only
# The connection set by Set-ConfluenceContext. The credential is kept separately, only in memory,
# and is never returned by Get-ConfluenceContext or written to any output stream.
$script:ConfluenceContext = $null
$script:ConfluenceCredential = $null
# Space key -> space ID lookups, cleared whenever the context changes
$script:ConfluenceSpaceIdCache = $null
# Nesting depth of New-ConfluenceContentTable while it renders nested tables
$script:ConfluenceTableNesting = 0
#endregion

#region module config, load telemetry and update check (never blocks import)
try {
    $CurrentConfig = Get-ModuleConfig -CommandPath $PSCommandPath -ErrorAction Stop
    Invoke-TelemetryCollection -ModuleName $CurrentConfig.ModuleName -ModuleVersion $CurrentConfig.ModuleVersion -CommandName 'Import-Module' -ExecutionID ([guid]::NewGuid().ToString()) -Stage 'Module-Load'
    if ($CurrentConfig.UpdateWarning -eq $true) {
        $null = Get-ModuleStatus -ShowMessage -ModuleName $CurrentConfig.ModuleName -ModulePath $CurrentConfig.ModulePath -CacheHours $CurrentConfig.UpdateCheckIntervalHours
    }
}
catch {
    Write-Warning "tcs.confluence configuration could not be loaded; defaults will be used. $($_.Exception.Message)"
}
#endregion

Export-ModuleMember -Function $Public.BaseName -Alias '*'
