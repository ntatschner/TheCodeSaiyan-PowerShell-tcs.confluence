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

Describe 'New-ConfluenceContentStatus' {
    It 'Creates the status macro' {
        New-ConfluenceContentStatus -Text 'Done' -Colour green |
            Should -BeExactly '<ac:structured-macro ac:name="status"><ac:parameter ac:name="colour">Green</ac:parameter><ac:parameter ac:name="title">Done</ac:parameter></ac:structured-macro>'
    }

    It 'Defaults to Grey, supports -Subtle and the -Color alias, and escapes the text' {
        $html = New-ConfluenceContentStatus -Text 'A&B'
        $html | Should -Match '>Grey<'
        $html | Should -Match '>A&amp;B<'
        New-ConfluenceContentStatus -Text 'x' -Color Blue -Subtle | Should -Match '<ac:parameter ac:name="subtle">true</ac:parameter>'
    }

    It 'Rejects unknown colours' {
        { New-ConfluenceContentStatus -Text 'x' -Colour Pink } | Should -Throw
    }
}

Describe 'New-ConfluenceContentExpand' {
    It 'Creates the expand macro with an escaped title and body' {
        New-ConfluenceContentExpand -Title 'More & more' -Content '<x>' |
            Should -BeExactly '<ac:structured-macro ac:name="expand"><ac:parameter ac:name="title">More &amp; more</ac:parameter><ac:rich-text-body><p>&lt;x&gt;</p></ac:rich-text-body></ac:structured-macro>'
    }

    It 'Inserts storage-format content with -Raw' {
        New-ConfluenceContentExpand -Content '<p>a</p>' -Raw |
            Should -BeExactly '<ac:structured-macro ac:name="expand"><ac:rich-text-body><p>a</p></ac:rich-text-body></ac:structured-macro>'
    }
}

Describe 'New-ConfluenceContentJiraIssue' {
    It 'Creates the jira macro for one issue' {
        New-ConfluenceContentJiraIssue -IssueKey ops-12 |
            Should -BeExactly '<ac:structured-macro ac:name="jira"><ac:parameter ac:name="key">OPS-12</ac:parameter><ac:parameter ac:name="showSummary">true</ac:parameter></ac:structured-macro>'
    }

    It 'Adds the server and serverId when given' {
        $html = New-ConfluenceContentJiraIssue -IssueKey OPS-1 -Server 'System JIRA' -ServerId 'abc-123' -ShowSummary $false
        $html | Should -Match '<ac:parameter ac:name="server">System JIRA</ac:parameter>'
        $html | Should -Match '<ac:parameter ac:name="serverId">abc-123</ac:parameter>'
        $html | Should -Match '<ac:parameter ac:name="showSummary">false</ac:parameter>'
    }

    It 'Creates a JQL table with escaped query, columns and maximum' {
        New-ConfluenceContentJiraIssue -JqlQuery 'project = OPS AND summary ~ "a<b"' -Columns key, summary -MaximumIssues 5 |
            Should -BeExactly '<ac:structured-macro ac:name="jira"><ac:parameter ac:name="jqlQuery">project = OPS AND summary ~ "a&lt;b"</ac:parameter><ac:parameter ac:name="columns">key,summary</ac:parameter><ac:parameter ac:name="maximumIssues">5</ac:parameter></ac:structured-macro>'
    }

    It 'Rejects keys that are not issue keys' {
        { New-ConfluenceContentJiraIssue -IssueKey 'OPS 1' } | Should -Throw
    }
}
