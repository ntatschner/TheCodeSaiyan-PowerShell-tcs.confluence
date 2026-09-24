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


Describe 'Get-ConfluencePageContent' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
    }

    BeforeEach {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"5","title":"Runbook","body":{"storage":{"value":"<p>x</p>"}}}' }
        }
    }

    It 'Gets a page body in storage format by default' {
        $result = Get-ConfluencePageContent -PageId 5
        $result.body.storage.value | Should -Be '<p>x</p>'
        $result.PSObject.Properties.Name | Should -Not -Contain 'Results'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/5?body-format=storage' }
    }

    It 'Requests the given body format' {
        $null = Get-ConfluencePageContent -PageId 5 -ContentType view
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*body-format=view' }
    }

    It 'Searches by space and exact title with the documented space-id parameter' {
        $null = Get-ConfluencePageContent -SpaceKey 42 -Title 'Runbook'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages?body-format=storage&limit=25&space-id=42&title=Runbook'
        }
    }

    It 'Sends exact titles with & # [ ] { } unchanged (URL-encoded)' {
        $null = Get-ConfluencePageContent -SpaceKey 42 -Title 'R&D #1 [draft] {x}'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Uri.OriginalString -like '*title=R%26D+%231+%5Bdraft%5D+%7Bx%7D' -and
            [System.Net.WebUtility]::UrlDecode($Uri.OriginalString) -like '*title=R&D #1 `[draft`] {x}'
        }
    }

    It 'Searches wildcard titles with CQL on content/search and requests the v1 expand=body.storage' {
        $null = Get-ConfluencePageContent -Title 'Run*'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $decoded = [System.Net.WebUtility]::UrlDecode($Uri.OriginalString)
            $decoded -like 'https://contoso.atlassian.net/wiki/rest/api/content/search?*' -and
            $decoded -like '*cql=type = page AND title ~ "Run*"*' -and
            $decoded -like '*expand=body.storage,*' -and $decoded -notlike '*body-format*'
        }
    }

    It 'Maps -ContentType to expand=body.<format> for wildcard searches' {
        $null = Get-ConfluencePageContent -Title 'Run*' -ContentType view
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            [System.Net.WebUtility]::UrlDecode($Uri.OriginalString) -like '*expand=body.view,*'
        }
    }

    It 'Writes each page to the pipeline' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1"},{"id":"2"}]}' }
        }
        $pages = @(Get-ConfluencePageContent -SpaceKey 42 -Title 'Runbook')
        $pages.id | Should -Be @('1', '2')
    }

    It 'Rejects unknown body formats' {
        { Get-ConfluencePageContent -PageId 5 -ContentType 'markdown' } | Should -Throw
    }
}
