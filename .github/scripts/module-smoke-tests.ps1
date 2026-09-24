[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Progress messages for the CI log; the script returns no output.')]
param()

$moduleName = 'tcs.confluence'
$repoRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$moduleDirectory = Join-Path -Path (Join-Path -Path $repoRoot -ChildPath 'modules') -ChildPath $moduleName
$moduleManifest = Join-Path -Path $moduleDirectory -ChildPath "$moduleName.psd1"

if (-not (Test-Path -Path $moduleManifest)) {
    throw "Module manifest not found at path: $moduleManifest"
}

# Keep the smoke test offline and away from the real user profile
$env:TCS_SKIP_UPDATE_CHECK = '1'
$env:TCS_TELEMETRY_OPTOUT = '1'
$env:TCS_CONFIG_ROOT = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "tcs-smoke-$([guid]::NewGuid().ToString('N'))"

try {
    Write-Host "Importing $moduleName from $moduleManifest" -ForegroundColor Cyan
    $importOutput = Import-Module -Name $moduleManifest -Force -ErrorAction Stop
    if ($null -ne $importOutput) { throw 'Importing the module wrote to the pipeline.' }

    $header = New-ConfluenceContentHeader -Header 'Smoke' -Level 2 -StringFormatting Bold
    if ($header -ne '<h2><strong>Smoke</strong></h2>') { throw "New-ConfluenceContentHeader returned '$header'." }

    $joined = Join-ConfluenceContent -ContentBlocks 'a', 'b' -Separator HorizontalRule
    if ($joined -ne 'a<hr />b') { throw "Join-ConfluenceContent returned '$joined'." }

    $link = New-ConfluenceContentLink -TextBlock 'see https://example.com'
    if ($link -ne "see <a href='https://example.com'>https://example.com</a>") { throw "New-ConfluenceContentLink returned '$link'." }

    $table = New-ConfluenceContentTable -TableData @([pscustomobject]@{ Name = 'x'; Value = 1 })
    if ($table -notmatch '^<table' -or $table -notmatch '</table>$') { throw 'New-ConfluenceContentTable did not return a table.' }

    $html = ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent '# Title'
    if ($html -ne '<h1>Title</h1>') { throw "ConvertTo-ConfluenceHTML returned '$html'." }

    $rows = @(Get-HtmlTableRowData -HtmlContent '<table><thead><tr><th>A</th><th>B</th></tr></thead><tbody><tr><td>1</td><td>2</td></tr></tbody></table>')
    if ($rows.Count -ne 1 -or $rows[0].B -ne '2') { throw 'Get-HtmlTableRowData did not parse the table.' }

    Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net/wiki' -Username 'smoke@example.com' -PersonalAccessToken 'smoke-token'
    $context = Get-ConfluenceContext
    if ($context.ConnectionURI -ne 'https://contoso.atlassian.net/wiki/api/v2') { throw "Unexpected ConnectionURI '$($context.ConnectionURI)'." }
    if (($context | ConvertTo-Json) -match 'smoke-token') { throw 'Get-ConfluenceContext exposed the token.' }

    $exported = @((Get-Module $moduleName).ExportedFunctions.Keys)
    $expected = @((Import-PowerShellDataFile -Path $moduleManifest).FunctionsToExport)
    $missing = $expected | Where-Object { $_ -notin $exported }
    if ($missing -or $exported.Count -ne $expected.Count) {
        throw "Exported functions do not match the manifest. Expected $($expected.Count), got $($exported.Count): $($exported -join ', ')"
    }

    Write-Host 'All smoke tests passed successfully.' -ForegroundColor Green
}
finally {
    Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $env:TCS_CONFIG_ROOT -Recurse -Force -ErrorAction SilentlyContinue
}
