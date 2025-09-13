#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\New-ConfluenceContentInternalLink.ps1"
    . "$PSScriptRoot\..\New-ConfluenceContentLink.ps1" # Dependency
}

Describe 'New-ConfluenceContentInternalLink' {
    InModuleScope -ModuleName 'tcs.confluence' {
        $script:ConfluenceContext = @{
            ConnectionBaseURL = 'https://mock.atlassian.net'
        }
    }
    Mock -CommandName Get-ConfluencePage -MockWith {
        param($PageId)
        return @{ Results = @{ _links = @{ webui = "/display/SPACE/$PageId" } } }
    } -Verifiable

    It 'should create a link from a PageId' {
        $html = New-ConfluenceContentInternalLink -PageId '12345' -LinkText 'My Link'
        $html | Should -Be '<a href=''https://mock.atlassian.net/wiki/display/SPACE/12345''>My Link</a>'
        Assert-MockCalled -CommandName Get-ConfluencePage -Scope It -Exactly 1
    }
}
