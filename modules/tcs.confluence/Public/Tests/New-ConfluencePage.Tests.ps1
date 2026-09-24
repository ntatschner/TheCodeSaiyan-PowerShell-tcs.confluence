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


Describe 'New-ConfluencePage' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
        $pageParams = @{ SpaceKey = '42'; ParentId = '100'; Title = 'Release notes'; Status = 'current'; Content = '<p>Hello</p>' }
        $conflict = [pscustomobject]@{ StatusCode = 400; StatusDescription = 'Bad Request'; Content = '{"errors":[{"title":"A page with this title already exists: A page already exists with the same TITLE in this space"}]}' }
    }

    Context 'Successful create' {
        BeforeEach {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"12345","title":"Release notes"}' }
            }
        }

        It 'POSTs the page and returns it' {
            $result = New-ConfluencePage @pageParams
            $result.id | Should -Be '12345'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $body = if ($Body) { [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json }
                $Method -eq 'POST' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages' -and
                $body.spaceId -eq '42' -and $body.parentId -eq '100' -and $body.title -eq 'Release notes' -and
                $body.body.representation -eq 'storage' -and $body.body.value -eq '<p>Hello</p>'
            }
        }

        It 'Sends nothing with -WhatIf' {
            New-ConfluencePage @pageParams -WhatIf | Should -BeNullOrEmpty
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
        }
    }

    Context 'Title conflict' {
        BeforeEach {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'POST' } { $conflict }
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'GET' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"7","title":"Release notes","parentId":"999","version":{"number":1}},{"id":"8","title":"Release notes","parentId":"100","version":{"number":4}}]}' }
            }
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'PUT' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"8","title":"Release notes","version":{"number":5}}' }
            }
        }

        It 'Reports the conflict without -Force and changes nothing' {
            { New-ConfluencePage @pageParams -ErrorAction Stop } | Should -Throw -ExpectedMessage '*already exists*-Force*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'PUT' }
        }

        It 'Updates the page with the same title under the same parent with -Force' {
            $result = New-ConfluencePage @pageParams -Force
            $result.id | Should -Be '8'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $body = if ($Body) { [System.Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json }
                $Method -eq 'PUT' -and $Uri.OriginalString -like '*/wiki/api/v2/pages/8' -and $body.version.number -eq 5 -and $body.body.value -eq '<p>Hello</p>'
            }
        }

        It 'Never overwrites a page with a different title' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'GET' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"9","title":"Release notes (old)","parentId":"100","version":{"number":1}}]}' }
            }
            { New-ConfluencePage @pageParams -Force -ErrorAction Stop } | Should -Throw -ExpectedMessage '*no page with exactly that title*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'PUT' }
        }
    }

    Context 'Other errors' {
        It 'Returns the page with a warning when Confluence created it despite an error' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'POST' } {
                [pscustomobject]@{ StatusCode = 500; StatusDescription = 'Internal Server Error'; Content = '' }
            }
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'GET' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"8","title":"Release notes","parentId":"100"}]}' }
            }
            $result = New-ConfluencePage @pageParams -WarningAction SilentlyContinue -WarningVariable pageWarning
            $result.id | Should -Be '8'
            $pageWarning | Should -Not -BeNullOrEmpty
        }

        It 'Reports the error when no page exists' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'POST' } {
                [pscustomobject]@{ StatusCode = 500; StatusDescription = 'Internal Server Error'; Content = '' }
            }
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'GET' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[]}' }
            }
            { New-ConfluencePage @pageParams -ErrorAction Stop } | Should -Throw -ExpectedMessage '*500*'
        }
    }
}
