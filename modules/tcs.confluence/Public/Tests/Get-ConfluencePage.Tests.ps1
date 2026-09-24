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


Describe 'Get-ConfluencePage' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
    }

    BeforeEach {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1","title":"Home"}]}' }
        }
    }

    It 'Gets a page by id' {
        $result = Get-ConfluencePage -PageId 123
        $result.Results[0].id | Should -Be '1'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/123' }
    }

    It 'Lists pages in a space with an exact title filter and limit' {
        $null = Get-ConfluencePage -SpaceKey 42 -Search 'Home' -ResultsLimit 10
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Uri.OriginalString -like 'https://contoso.atlassian.net/wiki/api/v2/pages?*' -and
            $Uri.OriginalString -like '*spaceId=42*' -and $Uri.OriginalString -like '*title=Home*' -and $Uri.OriginalString -like '*limit=10*'
        }
    }

    It 'Uses a CQL search for wildcard titles' {
        $null = Get-ConfluencePage -Search 'Ho*'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            [System.Net.WebUtility]::UrlDecode($Uri.OriginalString) -like '*/wiki/rest/api/content?*cql=title ~ "Ho*"*'
        }
    }

    It 'Reports request failures as errors' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 404; StatusDescription = 'Not Found'; Content = '' } }
        { Get-ConfluencePage -PageId 9 -ErrorAction Stop } | Should -Throw -ExpectedMessage '*404*'
    }
}
