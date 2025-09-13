#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\New-ConfluencePage.ps1"
    . "$PSScriptRoot\..\Update-ConfluencePage.ps1" # Dependency
}

Describe 'New-ConfluencePage' {
    # Mock the global context
    InModuleScope -ModuleName 'tcs.confluence' {
        $script:ConfluenceContext = @{
            ConnectionURI = 'https://mock.atlassian.net/wiki/api/v2'
            AuthorizationHeader = @{ Authorization = 'Basic mock' }
        }
    }

    Context 'When creating a new page successfully' {
        Mock -CommandName Invoke-WebRequest -MockWith {
            return @{
                StatusCode = 200
                Content = '{"id": "12345", "title": "New Page"}'
            } | ConvertTo-Json | ConvertFrom-Json -AsHashtable
        } -Verifiable

        It 'should call Invoke-WebRequest with POST and return the new page object' {
            $result = New-ConfluencePage -SpaceKey 'TEST' -ParentId '100' -Title 'New Page' -Status 'current' -Content '<p>Hello</p>'
            $result.id | Should -Be '12345'
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -ParameterFilter { $Method -eq 'Post' } -Exactly 1
        }
    }

    Context 'When page already exists and -Force is used' {
        # Simulate initial failure, then successful find and update
        $createFailedResponse = @{
            StatusCode = 400 # Or other error code
            Content = '{"Errors": {"title": "A page with this title already exists"}}'
        } | ConvertTo-Json | ConvertFrom-Json -AsHashtable

        Mock -CommandName Invoke-WebRequest -MockWith {
            param($Method)
            if ($Method -eq 'Post') { return $script:createFailedResponse }
            # This would be the PUT from Update-ConfluencePage
            if ($Method -eq 'Put') { return @{ StatusCode = 200; Content = '{"id": "existing-54321"}' } | ConvertTo-Json | ConvertFrom-Json -AsHashtable }
        } -Verifiable

        Mock -CommandName Get-ConfluencePage -MockWith {
            param($Search)
            if ($Search -eq 'Existing Page') {
                return @{ results = @( @{ id = 'existing-54321'; title = 'Existing Page'; version = @{ number = 1 } } ) }
            }
        } -Verifiable

        Mock -CommandName Update-ConfluencePage -MockWith {
            param($PageId)
            return @{ id = $PageId; title = 'Existing Page' }
        } -Verifiable

        It 'should attempt to create, fail, find the existing page, and update it' {
            $result = New-ConfluencePage -SpaceKey 'TEST' -ParentId '100' -Title 'Existing Page' -Status 'current' -Content '<p>Updated</p>' -Force
            
            $result.id | Should -Be 'existing-54321'
            Assert-MockCalled -CommandName Invoke-WebRequest -Scope It -ParameterFilter { $Method -eq 'Post' } -Exactly 1
            Assert-MockCalled -CommandName Get-ConfluencePage -Scope It -Exactly 1
            Assert-MockCalled -CommandName Update-ConfluencePage -Scope It -Exactly 1
        }
    }
}
