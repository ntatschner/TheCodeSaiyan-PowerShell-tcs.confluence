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

        It 'Sends a wildcard page title search to the documented CQL endpoint content/search' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'Test*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                (Get-DecodedUri $Uri) -eq 'https://contoso.atlassian.net/wiki/rest/api/content/search?cql=type = page AND title ~ "Test*"'
            }
        }

        It 'Sends a search on /content to content/search with an exact CQL match and escapes quotes' {
            $null = Invoke-ConfluenceRequest -Method GET -URIPath '/wiki/rest/api/content' -Search 'Say "hi"'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                (Get-DecodedUri $Uri) -eq 'https://contoso.atlassian.net/wiki/rest/api/content/search?cql=title = "Say \"hi\""'
            }
        }

        It 'Escapes a backslash before the quote so a title cannot break out of the CQL string' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'a\" OR space = "X*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                (Get-DecodedUri $Uri) -eq 'https://contoso.atlassian.net/wiki/rest/api/content/search?cql=type = page AND title ~ "a\\\" OR space = \"X*"'
            }
        }

        It 'Adds the space to the CQL query of a search (<Space>)' -ForEach @(
            @{ Space = 'DOCS'; Clause = 'space = "DOCS"' }
            @{ Space = '42'; Clause = 'space.id = 42' }
        ) {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'Run*' -Query @{ spaceKey = $Space }
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                (Get-DecodedUri $Uri) -eq "https://contoso.atlassian.net/wiki/rest/api/content/search?cql=type = page AND $Clause AND title ~ `"Run*`""
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
        BeforeEach {
            InModuleScope tcs.confluence { $script:ConfluenceSpaceIdCache = $null }
        }

        It 'Sends a numeric spaceKey as the documented space-id parameter without a lookup' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = '12345' }
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages?space-id=12345' }
        }

        It 'Sends spaceId as space-id' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceId = '7' }
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages?space-id=7' }
        }

        It 'Resolves a space key to its id once and caches it' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Uri.OriginalString -like '*/spaces?keys=DOCS*' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"98765","key":"DOCS"}]}' }
            }
            $query = @{ spaceKey = 'DOCS' }
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Query $query
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -Query $query
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages?space-id=98765' }
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*/spaces?keys=DOCS*' }
            $query.ContainsKey('spaceKey') | Should -BeTrue -Because 'the caller''s hashtable must not be changed'
        }

        It 'Reports an unknown space key and does not send an unfiltered request' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Uri.OriginalString -like '*/spaces?keys=*' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[]}' }
            }
            { Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = 'NOPE' } -ErrorAction Stop } | Should -Throw -ExpectedMessage '*NOPE*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly -ParameterFilter { $Uri.OriginalString -like '*/pages*' }
        }

        It 'Clears the cache when the context is set again' {
            InModuleScope tcs.confluence { $script:ConfluenceSpaceIdCache = New-Object -TypeName 'System.Collections.Generic.Dictionary[string,string]'; $script:ConfluenceSpaceIdCache['X'] = '1' }
            Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'secret-token-value'
            InModuleScope tcs.confluence { $script:ConfluenceSpaceIdCache } | Should -BeNullOrEmpty
        }
    }

    Context 'API version from the context' {
        AfterEach {
            Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'secret-token-value'
        }

        It 'Uses the v1 API for -Resource when the context was set with -ApiVersion v1' {
            Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'secret-token-value' -ApiVersion v1
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/rest/api/content' }
        }

        It 'Lets -ApiVersion override the context' {
            Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'secret-token-value' -ApiVersion v1
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -ApiVersion 2
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages' }
        }

        It 'Keeps Get-ConfluencePage on the v2 API it is written for' {
            Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'secret-token-value' -ApiVersion v1
            $null = Get-ConfluencePage -PageId 5
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/5' }
        }
    }

    Context 'Rate limiting' {
        BeforeEach {
            Mock -ModuleName tcs.confluence Start-Sleep { }
        }

        It 'Retries a 429 response after the Retry-After delay' {
            $script:calls = 0
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                $script:calls++
                if ($script:calls -eq 1) {
                    return [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = ''; Headers = @{ 'Retry-After' = @('3') } }
                }
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1"}]}' }
            }
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages -ErrorAction Stop
            @($result.Results).Count | Should -Be 1
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly
            Should -Invoke -ModuleName tcs.confluence Start-Sleep -Times 1 -Exactly -ParameterFilter { $Milliseconds -eq 3000 }
        }

        It 'Gives up after four retries and reports the 429' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                [pscustomobject]@{ StatusCode = 429; StatusDescription = 'Too Many Requests'; Content = '' }
            }
            { Invoke-ConfluenceRequest -Method GET -Resource pages -ErrorAction Stop } | Should -Throw -ExpectedMessage '*429*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 5 -Exactly
            Should -Invoke -ModuleName tcs.confluence Start-Sleep -Times 4 -Exactly
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

        It 'Stops at -MaxQueryPages and warns that more results exist' {
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages -MaxQueryPages 2 -WarningAction SilentlyContinue -WarningVariable pageWarning
            @($result.Results).Count | Should -Be 2
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 2 -Exactly
            "$pageWarning" | Should -BeLike '*more results*-All*'
        }

        It 'Does not warn when the last page was reached' {
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -MaxQueryPages 3 -WarningVariable pageWarning
            $pageWarning | Should -BeNullOrEmpty
        }

        It 'Reads every page with -All' {
            $result = Invoke-ConfluenceRequest -Method GET -Resource pages -MaxQueryPages 1 -All
            @($result.Results).id | Should -Be @('1', '2', '3')
        }

        It 'Resolves v1 next links against the /wiki context path' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest -ParameterFilter { $Uri.OriginalString -like '*/rest/api/content?limit=1' } {
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1"}],"_links":{"next":"/rest/api/content?limit=1&start=1"}}' }
            }
            $null = Invoke-ConfluenceRequest -Method GET -Resource pages -ApiVersion 1 -Query @{ limit = 1 } -MaxQueryPages 2 -WarningAction SilentlyContinue
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
