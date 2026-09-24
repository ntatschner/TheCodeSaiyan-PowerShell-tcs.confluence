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


Describe 'Invoke-ConfluenceRequest without a context' {
    BeforeAll {
        Import-Module -Name (Join-Path -Path $ModuleRoot -ChildPath 'tcs.confluence.psd1') -Force
        Mock -ModuleName tcs.confluence Invoke-WebRequest { }
    }

    It 'Reports an error and sends nothing' {
        $result = Invoke-ConfluenceRequest -Method GET -Resource pages -ErrorAction SilentlyContinue -ErrorVariable requestError
        $result | Should -BeNullOrEmpty
        "$requestError" | Should -Match 'Set-ConfluenceContext'
        Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
    }

    It 'Requires -URIPath or -Resource' {
        Invoke-ConfluenceRequest -Method GET -ErrorAction SilentlyContinue -ErrorVariable requestError
        "$requestError" | Should -Match 'URIPath or -Resource'
    }
}

Describe 'Invoke-ConfluenceRequest' {
    BeforeAll {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'secret-token-value'
        $expectedAuth = 'Basic ' + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('user@contoso.com:secret-token-value'))
        function Get-DecodedUri($Uri) { [System.Net.WebUtility]::UrlDecode($Uri.OriginalString) }
    }

    BeforeEach {
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1","title":"Page"}],"_links":{}}' }
        }
    }

    Context 'Endpoint construction' {
        It 'Builds <Expected> from <Description>' -ForEach @(
            @{ Description = 'v2 pages'; Params = @{ Resource = 'pages' }; Expected = 'https://contoso.atlassian.net/wiki/api/v2/pages' }
            @{ Description = 'v2 page by id'; Params = @{ Resource = 'pages'; Id = '12345' }; Expected = 'https://contoso.atlassian.net/wiki/api/v2/pages/12345' }
            @{ Description = 'v2 spaces'; Params = @{ Resource = 'spaces' }; Expected = 'https://contoso.atlassian.net/wiki/api/v2/spaces' }
            @{ Description = 'v1 pages'; Params = @{ Resource = 'pages'; ApiVersion = 1 }; Expected = 'https://contoso.atlassian.net/wiki/rest/api/content' }
            @{ Description = 'v1 spaces'; Params = @{ Resource = 'spaces'; ApiVersion = 1 }; Expected = 'https://contoso.atlassian.net/wiki/rest/api/space' }
            @{ Description = 'legacy /pages path'; Params = @{ URIPath = '/pages' }; Expected = 'https://contoso.atlassian.net/wiki/api/v2/pages' }
            @{ Description = 'explicit path'; Params = @{ URIPath = 'wiki/rest/api/content/5' }; Expected = 'https://contoso.atlassian.net/wiki/rest/api/content/5' }
        ) {
            $null = Invoke-ConfluenceRequest -Method GET @Params
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq $Expected }
        }

        It 'Switches a wildcard title search to the v1 CQL search' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'Test*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                (Get-DecodedUri $Uri) -eq 'https://contoso.atlassian.net/wiki/rest/api/content?cql=title ~ "Test*"'
            }
        }

        It 'Uses an exact CQL match without wildcards and escapes quotes' {
            $null = Invoke-ConfluenceRequest -Method GET -URIPath '/wiki/rest/api/content' -Search 'Say "hi"'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                (Get-DecodedUri $Uri) -eq 'https://contoso.atlassian.net/wiki/rest/api/content?cql=title = "Say \"hi\""'
            }
        }

        It 'URL-encodes query values' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ title = 'A&B C' }
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $Uri.OriginalString -like '*title=A%26B*C' -and (Get-DecodedUri $Uri) -like '*title=A&B C'
            }
        }

        It 'Does not duplicate the API path when the context URL contains it' {
            Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net/wiki/api/v2/' -Username 'user@contoso.com' -PersonalAccessToken 'secret-token-value'
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages' }
        }
    }

    Context 'spaceKey handling for v2 pages' {
        It 'Uses a numeric spaceKey as spaceId without a lookup' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = '12345' }
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*spaceId=12345*' }
        }

        It 'Resolves a space key to its id' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Uri.OriginalString -like '*/spaces?keys=DOCS*' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"98765","key":"DOCS"}]}' }
            }
            $query = @{ spaceKey = 'DOCS' }
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Query $query
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*/pages?spaceId=98765' }
            $query.ContainsKey('spaceKey') | Should -BeTrue -Because 'the caller''s hashtable must not be changed'
        }
    }

    Context 'Authentication' {
        It 'Sends a Basic Authorization header built from the stored credential' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Headers.Authorization -eq $expectedAuth }
        }

        It 'Never writes the token or the header to the verbose stream' {
            $verbose = Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = '1' } -Verbose 4>&1 | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] }
            $text = ($verbose | ForEach-Object { $_.Message }) -join "`n"
            $text | Should -Not -BeNullOrEmpty
            $text | Should -Not -Match 'secret-token-value'
            $text | Should -Not -Match ([regex]::Escape($expectedAuth.Substring(6)))
        }
    }

    Context 'Responses' {
        It 'Returns the results and marks multi-result responses' {
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages
            $result.MultiPage | Should -BeTrue
            @($result.Results).Count | Should -Be 1
            $result.Results[0].title | Should -Be 'Page'
        }

        It 'Wraps a single-object response in Results' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"id":"7","title":"One"}' } }
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages -Id 7
            $result.MultiPage | Should -BeFalse
            $result.Results[0].id | Should -Be '7'
        }

        It 'Accepts an empty body (204 No Content)' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 204; StatusDescription = 'No Content'; Content = '' } }
            $result = Invoke-ConfluenceRequest -Method DELETE -Resource pages -Id 7 -ErrorAction Stop
            @($result.Results).Count | Should -Be 0
        }

        It 'Reports HTTP errors with the status and Confluence error titles' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                [pscustomobject]@{ StatusCode = 400; StatusDescription = 'Bad Request'; Content = '{"errors":[{"status":400,"title":"A page with this title already exists"}]}' }
            }
            { Invoke-ConfluenceRequest -Method POST -Resource pages -Body '{}' -ErrorAction Stop } | Should -Throw -ExpectedMessage '*400*already exists*'
        }

        It 'Reports transport failures' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest { throw 'No such host is known.' }
            { Invoke-ConfluenceRequest -Method GET -Resource pages -ErrorAction Stop } | Should -Throw -ExpectedMessage '*No such host*'
        }

        It 'Sends the body as UTF-8 JSON' {
            $json = '{"title":"Gr' + [char]0x00F6 + 'sse"}'
            $null = Invoke-ConfluenceRequest -Method PUT -Resource pages -Id 1 -Body $json
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $Method -eq 'PUT' -and $ContentType -like 'application/json*' -and [System.Text.Encoding]::UTF8.GetString($Body) -eq $json
            }
        }

        It 'Rejects unsupported methods' {
            { Invoke-ConfluenceRequest -Method PATCH -Resource pages } | Should -Throw
        }
    }

    Context 'Pagination' {
        BeforeEach {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                if ($Uri.OriginalString -like '*cursor=abc*') {
                    return [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"2"}],"_links":{"next":"/wiki/api/v2/pages?cursor=def"}}' }
                }
                if ($Uri.OriginalString -like '*cursor=def*') {
                    return [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"3"}],"_links":{}}' }
                }
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1"}],"_links":{"next":"/wiki/api/v2/pages?cursor=abc"}}' }
            }
        }

        It 'Follows _links.next until there are no more pages' {
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages
            @($result.Results).id | Should -Be @('1', '2', '3')
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages?cursor=abc' }
        }

        It 'Stops at -MaxQueryPages' {
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages -MaxQueryPages 2
            @($result.Results).Count | Should -Be 2
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly
        }

        It 'Resolves v1 next links against the /wiki context path' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Uri.OriginalString -like '*/rest/api/content?limit=1' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1"}],"_links":{"next":"/rest/api/content?limit=1&start=1"}}' }
            }
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -ApiVersion 1 -Query @{ limit = 1 } -MaxQueryPages 2
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/rest/api/content?limit=1&start=1' }
        }

        It 'Does not follow a next link to another host' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1"}],"_links":{"next":"https://attacker.example/steal"}}' }
            }
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages -WarningAction SilentlyContinue -WarningVariable pageWarning
            @($result.Results).Count | Should -Be 1
            $pageWarning | Should -Not -BeNullOrEmpty
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
        }

        It 'Does not paginate non-GET requests' {
            $null = Invoke-ConfluenceRequest -Method POST -Resource pages -Body '{}'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
        }
    }
}
