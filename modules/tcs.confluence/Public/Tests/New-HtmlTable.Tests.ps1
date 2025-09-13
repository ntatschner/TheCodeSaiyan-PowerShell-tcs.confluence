#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\New-HtmlTable.ps1"
}

Describe 'New-HtmlTable' {
    $testData = @(
        [pscustomobject]@{ Region = 'North'; Sales = 100 },
        [pscustomobject]@{ Region = 'North'; Sales = 150 },
        [pscustomobject]@{ Region = 'South'; Sales = 200 }
    )

    Context 'CreateNew Parameter Set' {
        It 'should create a new table with merged rows' {
            $testData = @(
                [pscustomobject]@{ Region = 'North'; Sales = 100 },
                [pscustomobject]@{ Region = 'North'; Sales = 150 },
                [pscustomobject]@{ Region = 'South'; Sales = 200 }
            )
            $html = New-HtmlTable -InputObject $testData -MergeRows -MergeColumns 'Region'
            $html | Should -Match '<td rowspan="2">North</td>'
        }
    }

    Context 'MergeExisting Parameter Set' {
        It 'should merge new data into an existing table' {
            $testData = @(
                [pscustomobject]@{ Region = 'North'; Sales = 100 },
                [pscustomobject]@{ Region = 'North'; Sales = 150 },
                [pscustomobject]@{ Region = 'South'; Sales = 200 }
            )
            $existingTable = @"
<table>
<thead><tr><th>Region</th><th>Sales</th></tr></thead>
<tbody>
<tr><td>West</td><td>50</td></tr>
</tbody>
</table>
"@
            $html = New-HtmlTable -InputObject $testData -ExistingHtmlTable $existingTable -MergeWithExisting
            $html | Should -Match '<tr><td>West</td><td>50</td></tr>'
            $html | Should -Match '<tr><td>North</td><td>100</td></tr>'
        }
    }
}
