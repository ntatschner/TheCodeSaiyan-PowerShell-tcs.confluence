#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Remove-DuplicateConfluencePages.ps1"
}

Describe 'Remove-DuplicateConfluencePages' {
    Mock -CommandName Get-ConfluencePage -MockWith {
        return @{
            Results = @(
                @{ id = '1'; title = 'Duplicate'; parentId = '100'; version = @{ number = 1 }; lastModified = '2023-01-01' },
                @{ id = '2'; title = 'Duplicate'; parentId = '100'; version = @{ number = 2 }; lastModified = '2023-01-02' },
                @{ id = '3'; title = 'Duplicate'; parentId = '100'; version = @{ number = 3 }; lastModified = '2023-01-03' }
            )
        }
    } -Verifiable

    Mock -CommandName Invoke-ConfluenceRequest -MockWith { return $true } -Verifiable

    Context 'Deletion Logic' {
        It 'should keep the newest page and delete others when -KeepNewest is specified' {
            $keptPage = Remove-DuplicateConfluencePages -SpaceKey 'TEST' -ParentId '100' -PageTitle 'Duplicate' -KeepNewest
            
            $keptPage.id | Should -Be '3'
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $Method -eq 'DELETE' -and $Id -eq '1' } -Exactly 1
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $Method -eq 'DELETE' -and $Id -eq '2' } -Exactly 1
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $Method -eq 'DELETE' -and $Id -eq '3' } -Exactly 0
        }

        It 'should keep the oldest page and delete others when -KeepNewest is $false' {
            $keptPage = Remove-DuplicateConfluencePages -SpaceKey 'TEST' -ParentId '100' -PageTitle 'Duplicate' -KeepNewest:$false
            
            $keptPage.id | Should -Be '1'
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $Method -eq 'DELETE' -and $Id -eq '2' } -Exactly 1
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $Method -eq 'DELETE' -and $Id -eq '3' } -Exactly 1
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $Method -eq 'DELETE' -and $Id -eq '1' } -Exactly 0
        }
    }

    Context '-WhatIf support' {
        It 'should not call delete when -WhatIf is present' {
            Remove-DuplicateConfluencePages -SpaceKey 'TEST' -ParentId '100' -PageTitle 'Duplicate' -KeepNewest -WhatIf
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $Method -eq 'DELETE' } -Exactly 0
        }
    }
}
