@{
    RootModule           = 'tcs.confluence.psm1'
    ModuleVersion        = '0.1.1'
    GUID                 = '8e75db44-4f93-4ab4-b151-06ec760efb21'
    Author               = 'Nigel Tatschner'
    CompanyName          = 'TheCodeSaiyan'
    Copyright            = '(c) 2025-2026 Nigel Tatschner. All rights reserved.'
    Description          = 'Functions to work with Confluence Cloud: a REST client (pages and spaces) and builders for Confluence storage-format content such as headings, tables, code blocks, links, layouts and tables of contents.'
    CompatiblePSEditions = @('Desktop', 'Core')
    PowerShellVersion    = '5.1'
    RequiredModules      = @(
        @{ ModuleName = 'tcs.core'; ModuleVersion = '0.3.0' }
    )
    FunctionsToExport    = @(
        'ConvertTo-ConfluenceHTML',
        'Get-ConfluenceContext',
        'Get-ConfluencePage',
        'Get-ConfluencePageContent',
        'Get-ConfluenceSpace',
        'Get-HtmlTableRowData',
        'Invoke-ConfluenceRequest',
        'Join-ConfluenceContent',
        'New-ConfluenceContentCodeBlock',
        'New-ConfluenceContentDivider',
        'New-ConfluenceContentHeader',
        'New-ConfluenceContentInfo',
        'New-ConfluenceContentInternalLink',
        'New-ConfluenceContentLink',
        'New-ConfluenceContentTable',
        'New-ConfluenceContentTOC',
        'New-ConfluencePage',
        'New-ConfluencePageLayout',
        'New-HtmlTable',
        'Remove-ConfluencePage',
        'Remove-DuplicateConfluencePage',
        'Set-ConfluenceContext',
        'Update-ConfluencePage'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @(
        'Get-ConfluenceSpaces',
        'Remove-DuplicateConfluencePages'
    )
    PrivateData          = @{
        PSData = @{
            Tags         = @('Confluence', 'Atlassian', 'Documentation', 'Wiki', 'REST', 'API', 'PSEdition_Desktop', 'PSEdition_Core', 'Windows', 'Linux', 'MacOS')
            ProjectUri   = 'https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence'
            LicenseUri   = 'https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence/blob/main/LICENSE'
            ReleaseNotes = 'https://github.com/ntatschner/TheCodeSaiyan-PowerShell-tcs.confluence/blob/main/CHANGELOG.md'
        }
    }
}
