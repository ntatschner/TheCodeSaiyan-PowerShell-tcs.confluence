@{
    RootModule           = 'tcs.confluence.psm1'
    ModuleVersion        = '0.2.0'
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
        'Clear-ConfluenceContext',
        'ConvertTo-ConfluenceHTML',
        'Get-ConfluenceContext',
        'Get-ConfluencePage',
        'Get-ConfluencePageContent',
        'Get-ConfluenceSpace',
        'Invoke-ConfluenceRequest',
        'Join-ConfluenceContent',
        'New-ConfluenceContentCodeBlock',
        'New-ConfluenceContentDivider',
        'New-ConfluenceContentExpand',
        'New-ConfluenceContentHeader',
        'New-ConfluenceContentInfo',
        'New-ConfluenceContentInternalLink',
        'New-ConfluenceContentJiraIssue',
        'New-ConfluenceContentLink',
        'New-ConfluenceContentStatus',
        'New-ConfluenceContentTable',
        'New-ConfluenceContentTOC',
        'New-ConfluencePage',
        'New-ConfluencePageLayout',
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
