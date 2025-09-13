#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Get-ConfluencePageContent.ps1"
}

Describe 'Get-ConfluencePageContent' {
    It 'should call Invoke-ConfluenceRequest with -Resource pages, -Id, and body-format query' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return @{ Results = @( @{ id = 'mocked' } ) } } -Verifiable
        Get-ConfluencePageContent -PageId '123' -ContentType 'storage'
        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'pages' -and $Id -eq '123' -and $Query.'body-format' -eq 'storage'
        } -Exactly 1
    }

    It 'should call Invoke-ConfluenceRequest with search parameters' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return @{ Results = @( @{ id = 'mocked' } ) } } -Verifiable
        Get-ConfluencePageContent -Title 'My Page' -SpaceKey 'TEST' -ContentType 'view'
        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'pages' -and $Query.title -eq 'My Page' -and $Query.spaceKey -eq 'TEST' -and $Query.'body-format' -eq 'view'
        } -Exactly 1
    }
}
