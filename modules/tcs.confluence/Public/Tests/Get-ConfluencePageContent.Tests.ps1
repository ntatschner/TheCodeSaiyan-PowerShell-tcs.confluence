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
        $result.Results[0].body.storage.value | Should -Be '<p>x</p>'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/5?body-format=storage' }
    }

    It 'Requests the given body format' {
        $null = Get-ConfluencePageContent -PageId 5 -ContentType view
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*body-format=view' }
    }

    It 'Searches by space and exact title' {
        $null = Get-ConfluencePageContent -SpaceKey 42 -Title 'Runbook'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Uri.OriginalString -like '*/wiki/api/v2/pages?*' -and $Uri.OriginalString -like '*title=Runbook*' -and $Uri.OriginalString -like '*spaceId=42*'
        }
    }

    It 'Searches wildcard titles with CQL' {
        $null = Get-ConfluencePageContent -Title 'Run*'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            [System.Net.WebUtility]::UrlDecode($Uri.OriginalString) -like '*/wiki/rest/api/content?*cql=title ~ "Run*"*'
        }
    }

    It 'Rejects unknown body formats' {
        { Get-ConfluencePageContent -PageId 5 -ContentType 'markdown' } | Should -Throw
    }
}
