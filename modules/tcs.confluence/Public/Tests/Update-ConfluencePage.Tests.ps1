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


Describe 'Update-ConfluencePage' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
    }

    BeforeEach {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"12345","title":"Updated","version":{"number":3}}' }
        }
    }

    It 'PUTs the new version and returns the page' {
        $result = Update-ConfluencePage -PageId 12345 -Title 'Updated' -Content '<p>New</p>' -Version 3 -VersionMessage 'Nightly'
        $result.id | Should -Be '12345'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $body = [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json
            $Method -eq 'PUT' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/12345' -and
            $body.id -eq '12345' -and $body.status -eq 'current' -and $body.version.number -eq 3 -and $body.version.message -eq 'Nightly' -and
            $body.body.value -eq '<p>New</p>' -and $null -eq $body.spaceId
        }
    }

    It 'Includes the space id when given (-SpaceKey still works as an alias)' {
        $null = Update-ConfluencePage -PageId 1 -SpaceKey 42 -Title 't' -Content 'c' -Version 2
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            ([System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json).spaceId -eq '42'
        }
    }

    It 'Reads the current version and sends the next one when -Version is not given' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'GET' } {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"12345","title":"Old","version":{"number":7}}' }
        }
        $null = Update-ConfluencePage -PageId 12345 -Title 'Updated' -Content '<p>New</p>'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'GET' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/12345'
        }
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'PUT' -and ([System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json).version.number -eq 8
        }
    }

    It 'Sends nothing with -WhatIf, even without -Version' {
        Update-ConfluencePage -PageId 1 -Title 't' -Content 'c' -WhatIf | Should -BeNullOrEmpty
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
    }

    It 'Sends nothing with -WhatIf' {
        Update-ConfluencePage -PageId 1 -Title 't' -Content 'c' -Version 2 -WhatIf | Should -BeNullOrEmpty
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
    }

    It 'Reports API errors' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 409; StatusDescription = 'Conflict'; Content = '{"errors":[{"title":"Version must be incremented"}]}' } }
        { Update-ConfluencePage -PageId 1 -Title 't' -Content 'c' -Version 2 -ErrorAction Stop } | Should -Throw -ExpectedMessage '*Version must be incremented*'
    }
}
