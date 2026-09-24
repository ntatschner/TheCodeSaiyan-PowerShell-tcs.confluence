BeforeAll {
    $env:TCS_CONFIG_ROOT = Join-Path -Path $TestDrive -ChildPath 'config'
    $env:TCS_SKIP_UPDATE_CHECK = '1'
    $env:TCS_TELEMETRY_OPTOUT = '1'
    $ModuleRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
    Import-Module -Name (Join-Path -Path $ModuleRoot -ChildPath 'tcs.confluence.psd1') -Force
    Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
}

AfterAll {
    Remove-Module -Name tcs.confluence -Force -ErrorAction SilentlyContinue
}

Describe 'Invoke-ConfluenceHttpRequest (private)' {
    It 'Returns status, description and content of a response' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 201; StatusDescription = 'Created'; Content = '{"id":"1"}' } }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/wiki/api/v2/pages' -Method POST -Body '{}' }
        $result.StatusCode | Should -Be 201
        $result.StatusDescription | Should -Be 'Created'
        $result.Content | Should -Be '{"id":"1"}'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $UseBasicParsing -and $Headers.Accept -eq 'application/json' }
    }

    It 'Skips the HTTP error check on PowerShell 7' -Skip:($PSVersionTable.PSVersion.Major -lt 7) {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '' } }
        $null = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $SkipHttpErrorCheck }
    }

    It 'Turns an HTTP error thrown by Windows PowerShell 5.1 into a response' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $exception = New-Object -TypeName System.Exception -ArgumentList 'The remote server returned an error: (404) Not Found.'
            Add-Member -InputObject $exception -MemberType NoteProperty -Name Response -Value ([pscustomobject]@{ StatusCode = 404; StatusDescription = 'Not Found' })
            $record = New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $exception, 'WebCmdletWebResponseException', 'InvalidOperation', $null
            $record.ErrorDetails = New-Object -TypeName System.Management.Automation.ErrorDetails -ArgumentList '{"errors":[{"title":"missing"}]}'
            throw $record
        }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        $result.StatusCode | Should -Be 404
        $result.StatusDescription | Should -Be 'Not Found'
        $result.Content | Should -Be '{"errors":[{"title":"missing"}]}'
    }

    It 'Rethrows transport failures' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { throw 'No such host is known.' }
        { InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET } } | Should -Throw -ExpectedMessage '*No such host*'
    }

    It 'Throws when no context is set' {
        InModuleScope tcs.confluence {
            $saved = $script:ConfluenceCredential
            try {
                $script:ConfluenceCredential = $null
                { Get-ConfluenceAuthHeader } | Should -Throw -ExpectedMessage '*Set-ConfluenceContext*'
            }
            finally {
                $script:ConfluenceCredential = $saved
            }
        }
    }
}

Describe 'Get-HtmlFormatTag (private)' {
    It 'Returns nested opening and reversed closing tags' {
        InModuleScope tcs.confluence { Get-HtmlFormatTag -Format Bold, Strikethrough } | Should -BeExactly '<strong><s>'
        InModuleScope tcs.confluence { Get-HtmlFormatTag -Format Bold, Strikethrough -Close } | Should -BeExactly '</s></strong>'
    }

    It 'Returns an empty string for no formats' {
        InModuleScope tcs.confluence { Get-HtmlFormatTag -Format $null } | Should -BeExactly ''
    }
}
