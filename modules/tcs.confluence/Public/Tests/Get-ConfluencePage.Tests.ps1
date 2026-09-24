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
        $result.id | Should -Be '1'
        $result.PSObject.Properties.Name | Should -Not -Contain 'Results'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/123' }
    }

    It 'Lists pages in a space with an exact title filter and limit' {
        $null = Get-ConfluencePage -SpaceKey 42 -Search 'Home' -ResultsLimit 10
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Uri.OriginalString -like 'https://contoso.atlassian.net/wiki/api/v2/pages?*' -and
            $Uri.OriginalString -like '*space-id=42*' -and $Uri.OriginalString -like '*title=Home*' -and $Uri.OriginalString -like '*limit=10*'
        }
    }

    It 'Uses the documented CQL endpoint content/search for wildcard titles' {
        $null = Get-ConfluencePage -Search 'Ho*'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            [System.Net.WebUtility]::UrlDecode($Uri.OriginalString) -eq 'https://contoso.atlassian.net/wiki/rest/api/content/search?cql=type = page AND title ~ "Ho*"&limit=25'
        }
    }

    It 'Writes page objects, not a Results wrapper' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1","title":"A"},{"id":"2","title":"B"}]}' }
        }
        $pages = @(Get-ConfluencePage -SpaceKey 42)
        $pages.Count | Should -Be 2
        $pages.title | Should -Be @('A', 'B')
    }

    Context 'Paging' {
        BeforeEach {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                if ($Uri.OriginalString -like '*cursor=*') {
                    return [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"2"}],"_links":{}}' }
                }
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1"}],"_links":{"next":"/wiki/api/v2/pages?cursor=abc"}}' }
            }
        }

        It 'Warns when -MaxQueryPages leaves results behind' {
            $pages = @(Get-ConfluencePage -SpaceKey 42 -MaxQueryPages 1 -WarningAction SilentlyContinue -WarningVariable pageWarning)
            $pages.Count | Should -Be 1
            "$pageWarning" | Should -BeLike '*more results*'
        }

        It 'Reads every page with -All' {
            $pages = @(Get-ConfluencePage -SpaceKey 42 -MaxQueryPages 1 -All -WarningVariable pageWarning)
            $pages.id | Should -Be @('1', '2')
            $pageWarning | Should -BeNullOrEmpty
        }
    }

    It 'Pipes pages into Remove-ConfluencePage' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'GET' } {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"5","title":"x"},{"id":"6","title":"y"}]}' }
        }
        Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'DELETE' } {
            [pscustomobject]@{ StatusCode = 204; StatusDescription = 'No Content'; Content = '' }
        }
        Get-ConfluencePage -SpaceKey 42 | Remove-ConfluencePage -Confirm:$false -ErrorAction Stop
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'DELETE' -and $Uri.OriginalString -like '*/pages/5' }
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'DELETE' -and $Uri.OriginalString -like '*/pages/6' }
    }

    It 'Reports request failures as errors' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 404; StatusDescription = 'Not Found'; Content = '' } }
        { Get-ConfluencePage -PageId 9 -ErrorAction Stop } | Should -Throw -ExpectedMessage '*404*'
    }
}
