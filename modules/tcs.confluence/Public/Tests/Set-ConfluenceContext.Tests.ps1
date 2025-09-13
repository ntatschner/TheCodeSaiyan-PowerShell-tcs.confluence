#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Set-ConfluenceContext.ps1"
}

Describe 'Set-ConfluenceContext' {
    Context 'Basic functionality' {
        It 'should create a global ConfluenceContext variable' {
            Set-ConfluenceContext -ConfluenceUrl 'https://my.atlassian.net' -Username 'user' -PersonalAccessToken 'token'
            $global:ConfluenceContext | Should -Not -BeNull
        }

        It 'should correctly encode the authorization header' {
            $username = 'testuser'
            $pat = 'testtoken'
            $expectedAuth = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("${username}:${pat}"))
            
            Set-ConfluenceContext -ConfluenceUrl 'https://my.atlassian.net' -Username $username -PersonalAccessToken $pat
            
            $global:ConfluenceContext.AuthorizationHeader.Authorization | Should -Be $expectedAuth
        }

        It 'should normalize the ConfluenceUrl and build the ConnectionURI for v2' {
            Set-ConfluenceContext -ConfluenceUrl 'https://my.atlassian.net/wiki/' -Username 'user' -PersonalAccessToken 'token' -ApiVersion 'v2'
            $global:ConfluenceContext.ConnectionBaseURL | Should -Be 'https://my.atlassian.net/wiki'
            $global:ConfluenceContext.ConnectionURI | Should -Be 'https://my.atlassian.net/wiki/api/v2'
        }

        It 'should normalize a URL with existing API path and build ConnectionURI for v1' {
            Set-ConfluenceContext -ConfluenceUrl 'https://my.atlassian.net/wiki/rest/api' -Username 'user' -PersonalAccessToken 'token' -ApiVersion 'v1'
            $global:ConfluenceContext.ConnectionBaseURL | Should -Be 'https://my.atlassian.net'
            $global:ConfluenceContext.ConnectionURI | Should -Be 'https://my.atlassian.net/wiki/rest/api'
        }
    }
}
