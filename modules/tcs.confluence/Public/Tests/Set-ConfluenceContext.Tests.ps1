BeforeAll {
    $env:TCS_CONFIG_ROOT = Join-Path -Path $TestDrive -ChildPath 'config'
    $env:TCS_SKIP_UPDATE_CHECK = '1'
    $env:TCS_TELEMETRY_OPTOUT = '1'
    $ModuleRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
    Import-Module -Name (Join-Path -Path $ModuleRoot -ChildPath 'tcs.confluence.psd1') -Force
}

AfterAll {
    Remove-Module -Name tcs.confluence -Force -ErrorAction SilentlyContinue
}


Describe 'Set-ConfluenceContext and Get-ConfluenceContext' {
    BeforeEach {
        Import-Module -Name (Join-Path -Path $ModuleRoot -ChildPath 'tcs.confluence.psd1') -Force
    }

    It 'Returns nothing before a context is set' {
        Get-ConfluenceContext | Should -BeNullOrEmpty
    }

    It 'Normalises <Url> to <Base> and derives the <Version> endpoint' -ForEach @(
        @{ Url = 'https://contoso.atlassian.net'; Version = 'v2'; Base = 'https://contoso.atlassian.net'; Endpoint = 'https://contoso.atlassian.net/wiki/api/v2' }
        @{ Url = 'https://contoso.atlassian.net/wiki/'; Version = 'v2'; Base = 'https://contoso.atlassian.net'; Endpoint = 'https://contoso.atlassian.net/wiki/api/v2' }
        @{ Url = 'https://contoso.atlassian.net/wiki/rest/api'; Version = 'v1'; Base = 'https://contoso.atlassian.net'; Endpoint = 'https://contoso.atlassian.net/wiki/rest/api' }
        @{ Url = ' https://contoso.atlassian.net/wiki/api/v2/ '; Version = 'v1'; Base = 'https://contoso.atlassian.net'; Endpoint = 'https://contoso.atlassian.net/wiki/rest/api' }
    ) {
        Set-ConfluenceContext -ConfluenceUrl $Url.Trim() -Username 'user@contoso.com' -PersonalAccessToken 'token' -ApiVersion $Version
        $context = Get-ConfluenceContext
        $context.ConnectionBaseURL | Should -Be $Base
        $context.ConnectionURI | Should -Be $Endpoint
        $context.ApiVersion | Should -Be $Version
        $context.Username | Should -Be 'user@contoso.com'
        $context.HasCredential | Should -BeTrue
    }

    It 'Accepts a PSCredential' {
        $secure = New-Object -TypeName System.Security.SecureString
        'cred-token'.ToCharArray() | ForEach-Object { $secure.AppendChar($_) }
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList 'cred@contoso.com', $secure
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Credential $credential
        (Get-ConfluenceContext).Username | Should -Be 'cred@contoso.com'
        InModuleScope tcs.confluence { (Get-ConfluenceAuthHeader).Authorization } |
            Should -Be ('Basic ' + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('cred@contoso.com:cred-token')))
    }

    It 'Never exposes the token' {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'super-secret-token' -Verbose 4>&1 |
            ForEach-Object { "$_" } | Should -Not -Match 'super-secret-token'
        (Get-ConfluenceContext | ConvertTo-Json) | Should -Not -Match 'super-secret-token'
        Get-Variable -Name ConfluenceContext -Scope Global -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        InModuleScope tcs.confluence { $script:ConfluenceCredential.Password } | Should -BeOfType [System.Security.SecureString]
    }

    It 'Rejects non-https URLs' {
        { Set-ConfluenceContext -ConfluenceUrl 'http://contoso.atlassian.net' -Username 'u' -PersonalAccessToken 't' } | Should -Throw
    }

    It 'Changes nothing with -WhatIf' {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'u' -PersonalAccessToken 't' -WhatIf
        Get-ConfluenceContext | Should -BeNullOrEmpty
    }
}
