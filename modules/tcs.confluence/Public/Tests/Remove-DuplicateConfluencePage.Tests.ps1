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


Describe 'Remove-DuplicateConfluencePage' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
        $pages = '{"results":[' +
        '{"id":"1","title":"Report","parentId":"100","version":{"number":1,"createdAt":"2024-01-01T00:00:00Z"}},' +
        '{"id":"2","title":"Report","parentId":"100","version":{"number":3,"createdAt":"2024-03-01T00:00:00Z"}},' +
        '{"id":"3","title":"Report","parentId":"100","version":{"number":2,"createdAt":"2024-02-01T00:00:00Z"}},' +
        '{"id":"4","title":"Report","parentId":"200","version":{"number":9}},' +
        '{"id":"5","title":"Other","parentId":"100","version":{"number":1}}]}'
    }

    BeforeEach {
        Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'GET' } {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = $pages }
        }
        Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Method -eq 'DELETE' } {
            [pscustomobject]@{ StatusCode = 204; StatusDescription = 'No Content'; Content = '' }
        }
    }

    It 'Keeps the newest page and deletes the other duplicates under the same parent' {
        $kept = Remove-DuplicateConfluencePage -SpaceKey 42 -ParentId 100 -PageTitle 'Report' -Confirm:$false
        $kept.id | Should -Be '2'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly -ParameterFilter { $Method -eq 'DELETE' }
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'DELETE' -and $Uri.OriginalString -like '*/pages/1' }
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'DELETE' -and $Uri.OriginalString -like '*/pages/3' }
    }

    It 'Keeps the oldest page with -KeepNewest:$false' {
        $kept = Remove-DuplicateConfluencePage -SpaceKey 42 -ParentId 100 -PageTitle 'Report' -KeepNewest:$false -Confirm:$false
        $kept.id | Should -Be '1'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'DELETE' -and $Uri.OriginalString -like '*/pages/1' }
    }

    It 'Deletes nothing with -WhatIf' {
        $kept = Remove-DuplicateConfluencePage -SpaceKey 42 -ParentId 100 -PageTitle 'Report' -WhatIf
        $kept.id | Should -Be '2'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'DELETE' }
    }

    It 'Returns the single page and deletes nothing when there is no duplicate' {
        $kept = Remove-DuplicateConfluencePage -SpaceKey 42 -ParentId 100 -PageTitle 'Other' -Confirm:$false
        $kept.id | Should -Be '5'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'DELETE' }
    }

    It 'Has ConfirmImpact High and is still available as Remove-DuplicateConfluencePages' {
        (Get-Command Remove-DuplicateConfluencePage).ScriptBlock.Attributes.Where({ $_ -is [System.Management.Automation.CmdletBindingAttribute] }).ConfirmImpact | Should -Be 'High'
        (Get-Command Remove-DuplicateConfluencePages).ResolvedCommand.Name | Should -Be 'Remove-DuplicateConfluencePage'
    }
}
