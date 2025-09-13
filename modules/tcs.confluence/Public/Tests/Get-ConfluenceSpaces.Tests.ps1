#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Get-ConfluenceSpaces.ps1"
}

Describe 'Get-ConfluenceSpaces' {
    It 'should get a specific space by ID' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith {
            param($Resource, $Id)
            return @{ id = $Id; name = 'Specific Space' }
        } -Verifiable

        Get-ConfluenceSpaces -SpaceId '123'
        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'spaces' -and $Id -eq '123'
        } -Exactly 1
    }

    It 'should get all spaces and filter them with -Search' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith {
            return @{
                results = @(
                    @{ id = '1'; name = 'Alpha Space' };
                    @{ id = '2'; name = 'Beta Space' };
                    @{ id = '3'; name = 'Gamma Project' }
                )
            }
        } -Verifiable

        $result = Get-ConfluenceSpaces -Search '*Space'
        $result.Count | Should -Be 2
        ($result.name -contains 'Alpha Space') | Should -BeTrue
        ($result.name -contains 'Beta Space') | Should -BeTrue
    }
}
