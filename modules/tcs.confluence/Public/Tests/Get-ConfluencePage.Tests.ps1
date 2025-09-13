#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Get-ConfluencePage.ps1"
}

Describe 'Get-ConfluencePage' {
    It 'should call Invoke-ConfluenceRequest with -Resource pages and -Id when PageId is provided' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return @{ Results = @( @{ id = 'mocked' } ) } } -Verifiable
        Get-ConfluencePage -PageId '123'
        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'pages' -and $Id -eq '123'
        } -Exactly 1
    }

    It 'should call Invoke-ConfluenceRequest with -Resource pages and query for SpaceKey' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return @{ Results = @( @{ id = 'mocked' } ) } } -Verifiable
        Get-ConfluencePage -SpaceKey 'MYSPACE'
        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'pages' -and $Query.spaceKey -eq 'MYSPACE'
        } -Exactly 1
    }

    It 'should call Invoke-ConfluenceRequest with an exact title search in Query' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return @{ Results = @( @{ id = 'mocked' } ) } } -Verifiable
        Get-ConfluencePage -Search 'My Exact Page'
        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'pages' -and $Query.title -eq 'My Exact Page'
        } -Exactly 1
    }

    It 'should call Invoke-ConfluenceRequest with a wildcard search using -Search' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return @{ Results = @( @{ id = 'mocked' } ) } } -Verifiable
        Get-ConfluencePage -Search 'My Page*'
        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'pages' -and $Search -eq 'My Page*'
        } -Exactly 1
    }
}
