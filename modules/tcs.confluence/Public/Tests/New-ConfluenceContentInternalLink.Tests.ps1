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


Describe 'New-ConfluenceContentInternalLink' {
    Context 'By URL' {
        It 'Links to the page URL' {
            New-ConfluenceContentInternalLink -InternalLinkURL 'https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook' -LinkText 'Runbook' |
                Should -BeExactly "<a href='https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook/'>Runbook</a>"
        }

        It 'Links to a heading using the page title' {
            New-ConfluenceContentInternalLink -InternalLinkURL 'https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook' -PageTitle 'My Runbook' -HeadingLink 'Restart steps' |
                Should -Match "href='https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook/#MyRunbook-Restartsteps'"
        }
    }

    Context 'By page id' {
        BeforeAll {
            Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
        }

        BeforeEach {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"123","title":"Run Book","_links":{"webui":"/spaces/DOCS/pages/123/Run+Book"}}' }
            }
        }

        It 'Looks up the page URL' {
            New-ConfluenceContentInternalLink -PageId 123 -LinkText 'Go' |
                Should -BeExactly "<a href='https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Run+Book'>Go</a>"
        }

        It 'Uses the looked-up title for heading links' {
            New-ConfluenceContentInternalLink -PageId 123 -HeadingLink 'Step 1' |
                Should -Match "#RunBook-Step1'"
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
        }

        It 'Looks up the title from the page id in the URL' {
            New-ConfluenceContentInternalLink -InternalLinkURL 'https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Run+Book' -HeadingLink 'Step 1' |
                Should -Match "#RunBook-Step1'"
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*/pages/123' }
        }

        It 'Reports an error when the page cannot be found' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"123"}' } }
            { New-ConfluenceContentInternalLink -PageId 123 -ErrorAction Stop } | Should -Throw -ExpectedMessage '*web URL*'
        }
    }
}
