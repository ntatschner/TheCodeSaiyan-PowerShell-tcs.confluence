#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Invoke-ConfluenceRequest.ps1"
    . "$PSScriptRoot\..\Set-ConfluenceContext.ps1"
}

Describe 'Invoke-ConfluenceRequest' {
    # Mock the global context and the web request cmdlet
    $mockResponse = @{
        StatusCode = 200
        Content = '{"results": [{"id": "123", "title": "Test Page"}]}'
        RelationLink = @{}
    } | ConvertTo-Json | ConvertFrom-Json -AsHashtable
    
    InModuleScope -ModuleName 'tcs.confluence' {
        Mock -CommandName Invoke-WebRequest -MockWith {
            param($Uri)
            Write-Verbose "Mock Invoke-WebRequest called with URI: $Uri"
            return $script:mockResponse
        }

        Set-ConfluenceContext -ConfluenceUrl 'https://my.atlassian.net' -Username 'user' -PersonalAccessToken 'token'
    }

    Context 'URL Construction and Normalization' {
        It 'should use -Resource to build a v2 pages path' {
            Invoke-ConfluenceRequest -Method GET -Resource 'pages'
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -ParameterFilter { $Uri -like 'https://my.atlassian.net/wiki/api/v2/pages' }
        }

        It 'should use -Resource and -Id to build a specific v2 pages path' {
            Invoke-ConfluenceRequest -Method GET -Resource 'pages' -Id '12345'
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -ParameterFilter { $Uri -like 'https://my.atlassian.net/wiki/api/v2/pages/12345' }
        }

        It 'should switch to v1 content endpoint for wildcard search on v2 pages' {
            Invoke-ConfluenceRequest -Method GET -Resource 'pages' -Search 'Test*'
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -ParameterFilter { $Uri -like 'https://my.atlassian.net/wiki/rest/api/content?cql=title~%22Test*%22' }
        }
    }

    Context 'Query Parameter Handling' {
        Mock -CommandName Invoke-ConfluenceRequest -ModuleName 'tcs.confluence' -MockWith {
            param($URIPath, $Query)
            if ($URIPath -like '*/spaces' -and $Query.keys -eq 'MYSPACE') {
                return @{ Results = @( @{ id = '98765'; key = 'MYSPACE' } ) }
            }
            Invoke-ConfluenceRequest @PSBoundParameters
        } -Verifiable

        It 'should resolve spaceKey to spaceId for v2 pages requests' {
            Invoke-ConfluenceRequest -Method GET -Resource 'pages' -Query @{ spaceKey = 'MYSPACE' }
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $URIPath -like '*/spaces' } -Exactly 1
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -ParameterFilter { $Uri -like '*spaceId=98765*' }
        }

        It 'should treat numeric spaceKey as spaceId' {
            Invoke-ConfluenceRequest -Method GET -Resource 'pages' -Query @{ spaceKey = '12345' }
            Assert-MockCalled -CommandName Invoke-ConfluenceRequest -Scope It -ParameterFilter { $URIPath -like '*/spaces' } -Exactly 0
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -ParameterFilter { $Uri -like '*spaceId=12345*' }
        }
    }

    Context 'Pagination' {
        It 'should follow the "next" link for pagination' {
            $page1Response = @{
                StatusCode = 200
                Content = '{"results": [{"id": "1"}], "links": {"next": "/wiki/api/v2/pages?cursor=abc"}}'
            } | ConvertTo-Json | ConvertFrom-Json -AsHashtable
            $page2Response = @{
                StatusCode = 200
                Content = '{"results": [{"id": "2"}], "links": {}}'
            } | ConvertTo-Json | ConvertFrom-Json -AsHashtable

            Mock -CommandName Invoke-WebRequest -MockWith {
                param($Uri)
                if ($Uri -like '*cursor=abc*') { return $script:page2Response }
                return $script:page1Response
            } -Verifiable

            $result = Invoke-ConfluenceRequest -Method GET -Resource 'pages'
            $result.Results.Count | Should -Be 2
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -Exactly 2
        }
    }
}
