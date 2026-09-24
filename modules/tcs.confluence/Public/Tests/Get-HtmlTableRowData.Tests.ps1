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


Describe 'Get-HtmlTableRowData' {
    It 'Parses headers and Confluence macros' {
        $html = @"
<table>
 <thead><tr><th>Task</th><th>Assignee</th><th>Link</th><th>Page</th></tr></thead>
 <tbody>
   <tr>
     <td><ac:structured-macro ac:name="jira"><ac:parameter ac:name="key">PROJ-123</ac:parameter></ac:structured-macro></td>
     <td><ac:userlink ac:username="jdoe">John Doe</ac:userlink></td>
     <td><a href="https://example.com">Example &amp; Co</a></td>
     <td><ac:link><ri:page ri:content-title="Docs"/><ac:plain-text-link-body><![CDATA[Project Docs]]></ac:plain-text-link-body></ac:link></td>
   </tr>
 </tbody>
</table>
"@
        $result = @(Get-HtmlTableRowData -HtmlContent $html)
        $result.Count | Should -Be 1
        $result[0].Task | Should -Be 'PROJ-123'
        $result[0].Assignee | Should -Be 'John Doe'
        $result[0].Link | Should -Be 'Example & Co'
        $result[0].Page | Should -Be 'Project Docs'
    }

    It 'Uses ColumnN names with -NoHeader' {
        $html = '<table><tbody><tr><td>Data1</td><td>Data2</td></tr></tbody></table>'
        $result = @(Get-HtmlTableRowData -HtmlContent $html -NoHeader)
        $result[0].Column1 | Should -Be 'Data1'
        $result[0].Column2 | Should -Be 'Data2'
    }

    It 'Reads row headers into the named column' {
        $html = '<table><tbody><tr><th>Row A</th><td>1</td></tr><tr><th>Row B</th><td>2</td></tr></tbody></table>'
        $result = @(Get-HtmlTableRowData -HtmlContent $html -NoHeader -RowHeaderColumnName 'Key')
        $result.Count | Should -Be 2
        $result[1].Key | Should -Be 'Row B'
        $result[1].Column1 | Should -Be '2'
    }

    It 'Selects a table by index' {
        $html = '<table><tbody><tr><td>first</td></tr></tbody></table><table><tbody><tr><td>second</td></tr></tbody></table>'
        (Get-HtmlTableRowData -HtmlContent $html -TableIndex 1 -NoHeader).Column1 | Should -Be 'second'
    }

    It 'Keeps entities encoded with -DecodeHtmlEntities $false' {
        $html = '<table><tbody><tr><td>a &amp; b</td></tr></tbody></table>'
        (Get-HtmlTableRowData -HtmlContent $html -NoHeader -DecodeHtmlEntities $false).Column1 | Should -Be 'a &amp; b'
    }

    It 'Warns and returns nothing when there is no table' {
        $result = Get-HtmlTableRowData -HtmlContent '<p>none</p>' -WarningAction SilentlyContinue -WarningVariable tableWarning
        $result | Should -BeNullOrEmpty
        $tableWarning | Should -Not -BeNullOrEmpty
    }

    It 'Rejects an invalid row header column name before parsing' {
        { Get-HtmlTableRowData -HtmlContent '<table></table>' -RowHeaderColumnName '1 bad' } | Should -Throw
    }
}
