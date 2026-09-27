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
        Mock -ModuleName tcs.confluence Invoke-WebRequest { }
        InModuleScope tcs.confluence {
            $saved = $script:ConfluenceCredential
            try {
                $script:ConfluenceCredential = $null
                { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET } |
                    Should -Throw -ExpectedMessage 'No Confluence context is set. Run Set-ConfluenceContext first.'
            }
            finally {
                $script:ConfluenceCredential = $saved
            }
        }
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
    }

    It 'Builds the Authorization header per request and merges additional headers' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '' } }
        $null = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method PUT -AdditionalHeaders @{ 'X-Atlassian-Token' = 'no-check' } }
        $expected = 'Basic ' + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('user@contoso.com:token'))
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
            $Headers.Authorization -eq $expected -and $Headers.Accept -eq 'application/json' -and $Headers['X-Atlassian-Token'] -eq 'no-check'
        }
        InModuleScope tcs.confluence { @(Get-Variable -Scope Script | Where-Object { "$($_.Value)" -like 'Basic *' }) } | Should -BeNullOrEmpty
    }
}

Describe 'Invoke-ConfluenceHttpRequest retries (private)' {
    BeforeEach {
        # Invoke-WithRetry sleeps in the tcs.core module scope
        Mock -ModuleName tcs.core Start-Sleep { }
    }

    It 'Retries a 429 after the Retry-After delay in seconds' {
        $script:calls = 0
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $script:calls++
            if ($script:calls -eq 1) {
                return [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = ''; Headers = @{ 'Retry-After' = @('7') } }
            }
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"1"}' }
        }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        $result.StatusCode | Should -Be 200
        $result.Content | Should -Be '{"id":"1"}'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 7000 }
    }

    It 'Uses Retry-After even when it is shorter than the backoff' {
        $script:calls = 0
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $script:calls++
            if ($script:calls -le 3) {
                return [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = ''; Headers = @{ 'Retry-After' = @('1') } }
            }
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '' }
        }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        $result.StatusCode | Should -Be 200
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 4 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 3 -Exactly -ParameterFilter { $Milliseconds -eq 1000 }
    }

    It 'Waits until a Retry-After HTTP date' {
        $script:calls = 0
        $script:retryAt = [datetime]::UtcNow.AddSeconds(30).ToString('r', [System.Globalization.CultureInfo]::InvariantCulture)
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $script:calls++
            if ($script:calls -eq 1) {
                return [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = ''; Headers = @{ 'Retry-After' = @($script:retryAt) } }
            }
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '' }
        }
        $null = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -ge 28000 -and $Milliseconds -le 31000 }
    }

    It 'Caps the Retry-After delay at 60 seconds' {
        $script:calls = 0
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $script:calls++
            if ($script:calls -eq 1) {
                return [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = ''; Headers = @{ 'Retry-After' = @('3600') } }
            }
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '' }
        }
        $null = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 60000 }
    }

    It 'Gives up after four retries with 2, 4, 8 and 16 second waits and returns the 429' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = '{"message":"slow down"}' }
        }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        $result.StatusCode | Should -Be 429
        $result.StatusDescription | Should -Be 'Too Many Requests'
        $result.Content | Should -Be '{"message":"slow down"}'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 5 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 4 -Exactly
        foreach ($expectedDelay in 2000, 4000, 8000, 16000) {
            Should -Invoke -ModuleName tcs.core Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq $expectedDelay }
        }
    }

    It 'Retries a 429 thrown by Windows PowerShell 5.1' {
        $script:calls = 0
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $script:calls++
            if ($script:calls -eq 1) {
                $exception = New-Object -TypeName System.Exception -ArgumentList 'The remote server returned an error: (429) Too Many Requests.'
                $webResponse = [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Headers = @{ 'Retry-After' = '5' } }
                Add-Member -InputObject $exception -MemberType NoteProperty -Name Response -Value $webResponse
                throw (New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $exception, 'WebCmdletWebResponseException', 'InvalidOperation', $null)
            }
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '' }
        }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        $result.StatusCode | Should -Be 200
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 5000 }
    }

    It 'Returns a 429 at once with -MaxRetry 0' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = '' } }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET -MaxRetry 0 }
        $result.StatusCode | Should -Be 429
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 0 -Exactly
    }

    It 'Does not retry a <StatusCode> response' -ForEach @(
        @{ StatusCode = 500; StatusDescription = 'Internal Server Error' }
        @{ StatusCode = 503; StatusDescription = 'Service Unavailable' }
        @{ StatusCode = 404; StatusDescription = 'Not Found' }
    ) {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = $StatusCode; StatusDescription = $StatusDescription; Content = ''; Headers = @{ 'Retry-After' = @('1') } }
        }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        $result.StatusCode | Should -Be $StatusCode
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 0 -Exactly
    }

    It 'Does not retry a 500 thrown by Windows PowerShell 5.1' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $exception = New-Object -TypeName System.Exception -ArgumentList 'The remote server returned an error: (500) Internal Server Error.'
            Add-Member -InputObject $exception -MemberType NoteProperty -Name Response -Value ([pscustomobject]@{ StatusCode = 500; StatusDescription = 'Internal Server Error' })
            throw (New-Object -TypeName System.Management.Automation.ErrorRecord -ArgumentList $exception, 'WebCmdletWebResponseException', 'InvalidOperation', $null)
        }
        $result = InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET }
        $result.StatusCode | Should -Be 500
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 0 -Exactly
    }

    It 'Does not retry a transport failure' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { throw (New-Object -TypeName System.Net.WebException -ArgumentList 'No such host is known.', ([System.Net.WebExceptionStatus]::NameResolutionFailure)) }
        { InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET } } |
            Should -Throw -ExpectedMessage 'No such host is known.'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 0 -Exactly
    }

    It 'Does not retry a transport failure after a 429' {
        $script:calls = 0
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            $script:calls++
            if ($script:calls -eq 1) {
                return [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = '' }
            }
            throw 'The operation has timed out.'
        }
        { InModuleScope tcs.confluence { Invoke-ConfluenceHttpRequest -Uri 'https://contoso.atlassian.net/x' -Method GET } } |
            Should -Throw -ExpectedMessage 'The operation has timed out.'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly
        Should -Invoke -ModuleName tcs.core Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 2000 }
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
