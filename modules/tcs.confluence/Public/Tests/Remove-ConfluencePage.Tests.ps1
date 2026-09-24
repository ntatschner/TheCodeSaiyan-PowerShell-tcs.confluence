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


Describe 'Remove-ConfluencePage' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
    }

    BeforeEach {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 204; StatusDescription = 'No Content'; Content = '' } }
    }

    It 'Has ConfirmImpact High' {
        (Get-Command Remove-ConfluencePage).ScriptBlock.Attributes.Where({ $_ -is [System.Management.Automation.CmdletBindingAttribute] }).ConfirmImpact | Should -Be 'High'
    }

    It 'Sends DELETE for the page and writes nothing to the pipeline' {
        Remove-ConfluencePage -PageId 12345 -Confirm:$false -ErrorAction Stop | Should -BeNullOrEmpty
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'DELETE' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/12345'
        }
    }

    It 'Accepts pages from the pipeline' {
        @([pscustomobject]@{ id = '1' }, [pscustomobject]@{ id = '2' }) | Remove-ConfluencePage -Confirm:$false
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly
    }

    It 'Sends nothing with -WhatIf' {
        Remove-ConfluencePage -PageId 12345 -WhatIf
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
    }

    It 'Moves the page to the trash and then purges it with -Purge' {
        Remove-ConfluencePage -PageId 12345 -Purge -Confirm:$false -ErrorAction Stop
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'DELETE' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/12345'
        }
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'DELETE' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/12345?purge=true'
        }
    }

    It 'Still purges a page that is already in the trash' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Uri.OriginalString -notlike '*purge=true' } {
            [pscustomobject]@{ StatusCode = 404; StatusDescription = 'Not Found'; Content = '' }
        }
        Remove-ConfluencePage -PageId 12345 -Purge -Confirm:$false -ErrorAction Stop
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*purge=true' }
    }

    It 'Reports a failed purge' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 403; StatusDescription = 'Forbidden'; Content = '' } }
        { Remove-ConfluencePage -PageId 12345 -Purge -Confirm:$false -ErrorAction Stop } | Should -Throw -ExpectedMessage '*purge*12345*403*'
    }

    It 'Reports failures' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 404; StatusDescription = 'Not Found'; Content = '' } }
        { Remove-ConfluencePage -PageId 12345 -Confirm:$false -ErrorAction Stop } | Should -Throw -ExpectedMessage '*12345*404*'
    }
}
