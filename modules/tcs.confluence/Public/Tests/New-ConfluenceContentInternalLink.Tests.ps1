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
    Context 'By title (storage-format page link)' {
        It 'Creates an ac:link to the page in the given space' {
            New-ConfluenceContentInternalLink -PageTitle 'Runbook' -SpaceKey OPS -LinkText 'Run it' |
                Should -BeExactly '<ac:link><ri:page ri:content-title="Runbook" ri:space-key="OPS"/><ac:link-body>Run it</ac:link-body></ac:link>'
        }

        It 'Links to a heading with ac:anchor and escapes every value' {
            New-ConfluenceContentInternalLink -PageTitle 'R&D "plan"' -HeadingLink 'Step <1>' |
                Should -BeExactly '<ac:link ac:anchor="Step &lt;1&gt;"><ri:page ri:content-title="R&amp;D &quot;plan&quot;"/><ac:link-body>R&amp;D "plan"</ac:link-body></ac:link>'
        }
    }

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

        It 'Looks up the title and space key and creates an ac:link' {
            New-ConfluenceContentInternalLink -PageId 123 -LinkText 'Go' |
                Should -BeExactly '<ac:link><ri:page ri:content-title="Run Book" ri:space-key="DOCS"/><ac:link-body>Go</ac:link-body></ac:link>'
        }

        It 'Looks up the page URL with -AsUrl' {
            New-ConfluenceContentInternalLink -PageId 123 -LinkText 'Go' -AsUrl |
                Should -BeExactly "<a href='https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Run+Book'>Go</a>"
        }

        It 'Uses the looked-up title for heading links with -AsUrl' {
            New-ConfluenceContentInternalLink -PageId 123 -HeadingLink 'Step 1' -AsUrl |
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
            { New-ConfluenceContentInternalLink -PageId 123 -AsUrl -ErrorAction Stop } | Should -Throw -ExpectedMessage '*web URL*'
        }
    }
}
