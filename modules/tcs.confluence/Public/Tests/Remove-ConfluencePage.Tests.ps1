#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Public\Remove-ConfluencePage.ps1"
    . "$PSScriptRoot\..\Public\Set-ConfluenceContext.ps1"
}

Describe 'Remove-ConfluencePage' {
    BeforeEach {
        Set-ConfluenceContext -ConfluenceUrl "https://example.atlassian.net/wiki" -Username "user" -PersonalAccessToken "token"
    }

    It 'should call Invoke-ConfluenceRequest with DELETE method and correct PageId' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return $true } -Verifiable

        Remove-ConfluencePage -PageId '12345' -Confirm:$false

        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter {
            $Resource -eq 'pages/12345' -and $Method -eq 'DELETE'
        } -Exactly 1
    }

    It 'should not call Invoke-ConfluenceRequest when -WhatIf is specified' {
        Mock -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -MockWith { return $true } -Verifiable

        Remove-ConfluencePage -PageId '12345' -WhatIf

        Assert-MockCalled -ModuleName tcs.confluence -CommandName Invoke-ConfluenceRequest -Scope It -Exactly 0
    }
}
