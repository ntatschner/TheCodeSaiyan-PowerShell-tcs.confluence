BeforeDiscovery {
    $RepoRoot = Split-Path -Path $PSScriptRoot -Parent
    $ManifestPath = Join-Path -Path (Join-Path -Path $RepoRoot -ChildPath 'modules/tcs.confluence') -ChildPath 'tcs.confluence.psd1'
    $ExportedFunctions = @((Import-PowerShellDataFile -Path $ManifestPath).FunctionsToExport | ForEach-Object { @{ Name = $_ } })
}

BeforeAll {
    $env:TCS_CONFIG_ROOT = Join-Path -Path $TestDrive -ChildPath 'config'
    $env:TCS_SKIP_UPDATE_CHECK = '1'
    $env:TCS_TELEMETRY_OPTOUT = '1'
    $RepoRoot = Split-Path -Path $PSScriptRoot -Parent
    $ManifestPath = Join-Path -Path (Join-Path -Path $RepoRoot -ChildPath 'modules/tcs.confluence') -ChildPath 'tcs.confluence.psd1'
    Import-Module -Name $ManifestPath -Force
}

AfterAll {
    Remove-Module -Name tcs.confluence -Force -ErrorAction SilentlyContinue
}

# Coverage guard: every exported command must report telemetry through the tcs.core token API, so a
# new command cannot skip it or bring back the old hand-written telemetry block. The token API is used
# instead of Invoke-TcsCommand, which re-emits the body's errors through its own cmdlet and stops
# -WarningVariable, -InformationVariable (and -ErrorVariable with -ErrorAction SilentlyContinue) from
# collecting what the command writes.
Describe 'Telemetry coverage for <Name>' -ForEach $ExportedFunctions {
    It 'Starts and completes a tcs.core telemetry token' {
        $definition = (Get-Command -Name $Name -Module tcs.confluence).Definition
        $definition | Should -Match '\$telemetry\s*=\s*Start-TcsTelemetry'
        $definition | Should -Match 'Complete-TcsTelemetry\s+-Token\s+\$telemetry\s+-ErrorRecord'
        $definition | Should -Match 'Complete-TcsTelemetry\s+-Token\s+\$telemetry\s*[\r\n]'
        $definition | Should -Not -Match 'Invoke-TelemetryCollection'
        $definition | Should -Not -Match 'Invoke-TcsCommand'
    }
}

Describe 'Telemetry wiring' {
    It 'Public command files call neither Invoke-TelemetryCollection nor Invoke-TcsCommand' {
        $publicFolder = Join-Path -Path (Join-Path -Path $RepoRoot -ChildPath 'modules/tcs.confluence') -ChildPath 'Public'
        $offenders = @(Get-ChildItem -Path $publicFolder -Filter '*.ps1' -Recurse |
                Where-Object { $_.Name -notlike '*.Tests.ps1' } |
                Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match 'Invoke-TelemetryCollection|Invoke-TcsCommand' } |
                ForEach-Object { $_.Name })
        $offenders | Should -BeNullOrEmpty
    }
}

Describe 'Telemetry behaviour' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
    }

    BeforeEach {
        Mock -ModuleName tcs.core Invoke-TelemetryCollection { }
    }

    It 'New-ConfluenceContentHeader sends Start and a successful End' {
        New-ConfluenceContentHeader -Header 'Title' -Level 1 | Should -BeExactly '<h1>Title</h1>'
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'New-ConfluenceContentHeader' -and $ModuleName -eq 'tcs.confluence' -and $Stage -eq 'Start' -and $ClearTimer
        }
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'New-ConfluenceContentHeader' -and $Stage -eq 'End' -and -not $Failed
        }
    }

    It 'Get-ConfluenceSpace sends a successful End after a request' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1","key":"ENG","name":"Engineering"}]}' }
        }
        @(Get-ConfluenceSpace).Count | Should -Be 1
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Get-ConfluenceSpace' -and $Stage -eq 'Start'
        }
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Get-ConfluenceSpace' -and $Stage -eq 'End' -and -not $Failed
        }
    }

    It 'Get-ConfluenceSpace is the only command reported, not the Invoke-ConfluenceRequest it calls' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[]}' }
        }
        $null = Get-ConfluenceSpace
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 0 -Exactly -ParameterFilter {
            $CommandName -eq 'Invoke-ConfluenceRequest'
        }
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 2 -Exactly
    }

    It 'Get-ConfluenceSpace sends a failed End and still writes its error when the request fails' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 500; StatusDescription = 'Server Error'; Content = '' }
        }
        $result = Get-ConfluenceSpace -ErrorAction SilentlyContinue -ErrorVariable spaceErrors
        $result | Should -BeNullOrEmpty
        "$spaceErrors" | Should -BeLike '*Failed to retrieve space information*'
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Get-ConfluenceSpace' -and $Stage -eq 'End' -and $Failed -and $null -ne $Exception
        }
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Get-ConfluenceSpace' -and $Stage -eq 'End'
        }
    }

    It 'Remove-ConfluencePage times the whole pipeline once' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 204; StatusDescription = 'No Content'; Content = '' } }
        @([pscustomobject]@{ id = '1' }, [pscustomobject]@{ id = '2' }) | Remove-ConfluencePage -Confirm:$false | Should -BeNullOrEmpty
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Remove-ConfluencePage' -and $Stage -eq 'Start'
        }
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Remove-ConfluencePage' -and $Stage -eq 'End' -and -not $Failed
        }
    }

    It 'Remove-ConfluencePage sends one failed End and rethrows on a terminating error' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 404; StatusDescription = 'Not Found'; Content = '' } }
        { Remove-ConfluencePage -PageId 12345 -Confirm:$false -ErrorAction Stop } | Should -Throw -ExpectedMessage '*12345*404*'
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Remove-ConfluencePage' -and $Stage -eq 'End' -and $Failed
        }
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Remove-ConfluencePage' -and $Stage -eq 'End'
        }
    }

    It 'Remove-ConfluencePage still sends telemetry with -WhatIf and sends no request' {
        Mock -ModuleName tcs.confluence Invoke-WebRequest { }
        Remove-ConfluencePage -PageId 12345 -WhatIf
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
        Should -Invoke -ModuleName tcs.core Invoke-TelemetryCollection -Times 1 -Exactly -ParameterFilter {
            $CommandName -eq 'Remove-ConfluencePage' -and $Stage -eq 'End' -and -not $Failed
        }
    }
}
