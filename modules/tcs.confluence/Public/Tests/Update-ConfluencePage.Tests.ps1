#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Update-ConfluencePage.ps1"
}

Describe 'Update-ConfluencePage' {
    Mock -CommandName Invoke-ConfluenceRequest -MockWith {
        param($Body)
        # Return a simple object that mimics the structure
        return @{
            Results = @(
                (ConvertFrom-Json $Body)
            )
        }
    } -Verifiable

    It 'should call Invoke-ConfluenceRequest with PUT and a correctly structured body' {
        $pageId = '12345'
        $version = 5
        $title = 'Updated Title'
        $content = '<p>Updated content</p>'

        $result = Update-ConfluencePage -PageId $pageId -Title $title -Content $content -Version $version

        $result.id | Should -Be $pageId
        $result.version.number | Should -Be $version

        Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Method -eq 'PUT' -and
            $Resource -eq 'pages' -and
            $Id -eq $pageId -and
            ($Body | ConvertFrom-Json).version.number -eq $version
        } -Exactly 1
    }
}
