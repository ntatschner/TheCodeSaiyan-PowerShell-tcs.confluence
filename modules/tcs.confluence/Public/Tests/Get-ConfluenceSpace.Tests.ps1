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


Describe 'Get-ConfluenceSpace' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
    }

    BeforeEach {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1","key":"ENG","name":"Engineering"},{"id":"2","key":"HR","name":"People"}]}' }
        }
    }

    It 'Lists all spaces' {
        $result = @(Get-ConfluenceSpace)
        $result.Count | Should -Be 2
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/spaces?limit=50' }
    }

    It 'Filters by name with wildcards' {
        $result = @(Get-ConfluenceSpace -Search 'eng*')
        $result.Count | Should -Be 1
        $result[0].key | Should -Be 'ENG'
    }

    It 'Gets one space by id' {
        $null = Get-ConfluenceSpace -SpaceId 2
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/spaces/2' }
    }

    It 'Is still available under the old name Get-ConfluenceSpaces' {
        (Get-Command -Name Get-ConfluenceSpaces).ResolvedCommand.Name | Should -Be 'Get-ConfluenceSpace'
        @(Get-ConfluenceSpaces).Count | Should -Be 2
    }
}
