#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\New-ConfluenceContentTable.ps1"
    . "$PSScriptRoot\..\New-ConfluenceContentLink.ps1" # Dependency
}

Describe 'New-ConfluenceContentTable' {
    It 'should create a simple table from an array of PSCustomObjects' {
        $data = @(
            [pscustomobject]@{ Name = 'Item 1'; Value = 100 },
            [pscustomobject]@{ Name = 'Item 2'; Value = 200 }
        )
        $html = New-ConfluenceContentTable -TableData $data

        $html | Should -Match '<thead><tr><th.*?>Name</th><th.*?>Value</th></tr></thead>'
        $html | Should -Match '<tbody>'
        $html | Should -Match '<tr><td.*?><span>Item 1</span></td><td.*?>100</td></tr>'
        $html | Should -Match '<tr><td.*?><span>Item 2</span></td><td.*?>200</td></tr>'
        $html | Should -Match '</tbody></table>'
    }
}
