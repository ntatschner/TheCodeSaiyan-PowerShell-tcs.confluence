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


Describe 'New-HtmlTable' {
    BeforeAll {
        $testData = @(
            [pscustomobject]@{ Region = 'North'; Sales = 100 }
            [pscustomobject]@{ Region = 'North'; Sales = 150 }
            [pscustomobject]@{ Region = 'South'; Sales = 200 }
        )
        # Compare markup without line breaks
        function ConvertTo-FlatHtml([string]$Html) { ($Html -replace "`r?`n", '') }
    }

    Context 'CreateNew' {
        It 'Creates a header and one row per object' {
            $html = ConvertTo-FlatHtml (New-HtmlTable -InputObject $testData)
            $html | Should -Match '<thead><tr><th>Region</th><th>Sales</th></tr></thead>'
            $html | Should -Match '<tr><td>North</td><td>100</td></tr>'
            ([regex]::Matches($html, '<tr>')).Count | Should -Be 4
        }

        It 'Accepts pipeline input' {
            $html = ConvertTo-FlatHtml ($testData | New-HtmlTable)
            ([regex]::Matches($html, '<tr>')).Count | Should -Be 4
        }

        It 'Uses only the requested properties, in order' {
            $html = ConvertTo-FlatHtml (New-HtmlTable -InputObject $testData -Properties Sales)
            $html | Should -Match '<thead><tr><th>Sales</th></tr></thead>'
            $html | Should -Not -Match 'North'
        }

        It 'Adds caption, class and styles' {
            $html = ConvertTo-FlatHtml (New-HtmlTable -InputObject $testData -Title 'Q1 <sales>' -CssClass 'report' -Style @{ width = '100%' } -HeaderStyle @{ color = 'red' } -CellStyle @{ padding = '2px' })
            $html | Should -Match '<table style="width: 100%;" class="report">'
            $html | Should -Match '<caption>Q1 &lt;sales&gt;</caption>'
            $html | Should -Match '<th style="color: red;">Region</th>'
            $html | Should -Match '<td style="padding: 2px;">North</td>'
        }

        It 'Merges equal consecutive values with rowspan' {
            $html = ConvertTo-FlatHtml (New-HtmlTable -InputObject $testData -MergeRows -MergeColumns 'Region')
            $html | Should -Match '<td rowspan="2">North</td><td>100</td>'
            $html | Should -Match '<tr><td>150</td></tr>'
        }

        It 'HTML-encodes values unless -PreserveHtml is given' {
            $data = [pscustomobject]@{ Value = '<b>x</b>' }
            ConvertTo-FlatHtml (New-HtmlTable -InputObject $data) | Should -Match '<td>&lt;b&gt;x&lt;/b&gt;</td>'
            ConvertTo-FlatHtml (New-HtmlTable -InputObject $data -PreserveHtml) | Should -Match '<td><b>x</b></td>'
        }

        It 'Shows -NullDisplay or &nbsp; for empty values' {
            $data = [pscustomobject]@{ A = $null }
            ConvertTo-FlatHtml (New-HtmlTable -InputObject $data -NullDisplay 'n/a') | Should -Match '<td>n/a</td>'
            ConvertTo-FlatHtml (New-HtmlTable -InputObject $data -NullDisplay '<i>-</i>' -PreserveHtml) | Should -Match '<td><i>-</i></td>'
            ConvertTo-FlatHtml (New-HtmlTable -InputObject $data -UseNbspForEmpty) | Should -Match '<td>&nbsp;</td>'
        }

        It 'Warns and ignores -MergeColumns without -MergeRows' {
            $html = New-HtmlTable -InputObject $testData -MergeColumns 'Region' -WarningVariable tableWarning -WarningAction SilentlyContinue
            $tableWarning | Should -Not -BeNullOrEmpty
            ConvertTo-FlatHtml $html | Should -Not -Match 'rowspan'
        }

        It 'Reports an error for merge columns that do not exist' {
            $html = New-HtmlTable -InputObject $testData -MergeRows -MergeColumns 'Missing' -ErrorAction SilentlyContinue -ErrorVariable tableError
            $html | Should -BeNullOrEmpty
            $tableError | Should -Not -BeNullOrEmpty
        }
    }

    Context 'MergeExisting' {
        BeforeAll {
            $existingTable = @"
<table>
<thead><tr><th>Region</th><th>Sales</th></tr></thead>
<tbody>
<tr><td>West</td><td>50</td></tr>
</tbody>
</table>
"@
        }

        It 'Appends the new rows after the existing rows' {
            $html = ConvertTo-FlatHtml (New-HtmlTable -InputObject $testData -ExistingHtmlTable $existingTable -MergeWithExisting)
            $html | Should -Match '<tbody><tr><td>West</td><td>50</td></tr><tr><td>North</td><td>100</td></tr>'
            ([regex]::Matches($html, '<tr>')).Count | Should -Be 5
        }

        It 'Applies -CellStyle to existing and new cells' {
            $html = ConvertTo-FlatHtml (New-HtmlTable -InputObject $testData -ExistingHtmlTable $existingTable -MergeWithExisting -CellStyle @{ color = 'blue' })
            $html | Should -Match '<td style="color: blue;">West</td>'
            $html | Should -Match '<td style="color: blue;">North</td>'
        }

        It 'Reports an error when the existing table has no header' {
            $result = New-HtmlTable -InputObject $testData -ExistingHtmlTable '<table><tbody></tbody></table>' -MergeWithExisting -ErrorAction SilentlyContinue -ErrorVariable mergeError
            $result | Should -BeNullOrEmpty
            $mergeError | Should -Not -BeNullOrEmpty
        }
    }
}
