#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\Get-HtmlTableRowData.ps1"
}

Describe 'Get-HtmlTableRowData' {
    It 'should parse a basic table with headers and macros' {
        $htmlWithMacros = @"
    <table>
     <thead><tr><th>Task</th><th>Assignee</th></tr></thead>
     <tbody>
       <tr>
         <td><ac:structured-macro ac:name="jira"><ac:parameter ac:name="key">PROJ-123</ac:parameter></ac:structured-macro></td>
         <td><ac:userlink ac:username="jdoe">John Doe</ac:userlink></td>
       </tr>
     </tbody>
    </table>
"@
        $result = Get-HtmlTableRowData -HtmlContent $htmlWithMacros
        $result.Count | Should -Be 1
        $result[0].Task | Should -Be 'PROJ-123'
        $result[0].Assignee | Should -Be 'jdoe'
    }

    It 'should handle tables without headers using -NoHeader' {
        $htmlNoHeader = @"
        <table>
         <tbody>
           <tr><td>Data1</td><td>Data2</td></tr>
         </tbody>
        </table>
"@
        $result = Get-HtmlTableRowData -HtmlContent $htmlNoHeader -NoHeader
        $result.Count | Should -Be 1
        $result[0].Column1 | Should -Be 'Data1'
        $result[0].Column2 | Should -Be 'Data2'
    }
}
