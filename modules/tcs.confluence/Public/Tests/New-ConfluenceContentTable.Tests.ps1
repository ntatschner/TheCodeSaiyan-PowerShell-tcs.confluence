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


Describe 'New-ConfluenceContentTable' {
    BeforeAll {
        $rows = @(
            [pscustomobject]@{ Name = 'alpha'; Count = 1 }
            [pscustomobject]@{ Name = 'beta'; Count = 2 }
        )
    }

    It 'Creates a header row from the property names' {
        $html = New-ConfluenceContentTable -TableData $rows
        $html | Should -Match "^<table class='Wrapped' style='Width: 100%'><thead><tr>"
        $html | Should -Match "<th style='text-align: Left;' scope='col'>Name</th><th style='text-align: Left;' scope='col'>Count</th>"
        $html | Should -Match '</tbody></table>$'
        ([regex]::Matches($html, '<tr>')).Count | Should -Be 3
    }

    It 'Keeps numeric values in the first column' {
        $html = New-ConfluenceContentTable -TableData @([pscustomobject]@{ Id = 42; Name = 'x' })
        $html | Should -Match "<td style='text-align: Left;'><span>42</span></td>"
    }

    It 'Closes first-cell formatting tags in the right order' {
        $html = New-ConfluenceContentTable -TableData $rows -FirstCellStringFormatting Bold, Italic
        $html | Should -Match '<span><strong><em>alpha</em></strong></span>'
    }

    It 'Nests header formatting tags correctly' {
        $html = New-ConfluenceContentTable -TableData $rows -HeaderStringFormatting Bold, Underline
        $html | Should -Match '<strong><u>Name</u></strong>'
    }

    It 'Applies cell formatting to the other columns' {
        $html = New-ConfluenceContentTable -TableData $rows -CellStringFormatting Italic -CellAlignmentFormatting Right
        $html | Should -Match "<td style='text-align: Right;'><em>1</em></td>"
    }

    It 'Wraps the first cell in a heading when requested' {
        New-ConfluenceContentTable -TableData $rows -FirstCellHeaderFormat 3 | Should -Match '<h3>alpha</h3>'
    }

    It 'Renders row headers with -VerticalHeader' {
        New-ConfluenceContentTable -TableData $rows -VerticalHeader | Should -Match "<th scope='row' style='text-align: Left;'><span>alpha</span></th>"
    }

    It 'Omits the header with -NoHeader' {
        New-ConfluenceContentTable -TableData $rows -NoHeader | Should -Not -Match '<thead>'
    }

    It 'Renders dates as text, not as nested tables' {
        $html = New-ConfluenceContentTable -TableData @([pscustomobject]@{ Name = 'x'; When = [datetime]'2020-01-02' })
        ([regex]::Matches($html, '<table')).Count | Should -Be 1
    }

    It 'Renders nested objects as nested tables' {
        $html = New-ConfluenceContentTable -TableData @([pscustomobject]@{ Name = 'x'; Detail = [pscustomobject]@{ A = 1; B = 2 } })
        ([regex]::Matches($html, '<table')).Count | Should -Be 2
    }

    It 'Converts URLs in data cells to links' {
        $html = New-ConfluenceContentTable -TableData @([pscustomobject]@{ Name = 'x'; Link = 'https://example.com' })
        $html | Should -Match "<a href='https://example.com'>https://example.com</a>"
    }

    It 'Returns an empty string, and only that, for an empty collection' {
        $result = @(New-ConfluenceContentTable -TableData @())
        $result.Count | Should -Be 1
        $result[0] | Should -BeExactly ''
    }

    It 'Reports an error and returns nothing for rows with different columns' {
        $mixed = @([pscustomobject]@{ A = 1 }, [pscustomobject]@{ B = 1 })
        $result = New-ConfluenceContentTable -TableData $mixed -ErrorAction SilentlyContinue -ErrorVariable tableError
        $result | Should -BeNullOrEmpty
        $tableError | Should -Not -BeNullOrEmpty
    }

    It 'Validates the table style' {
        { New-ConfluenceContentTable -TableData $rows -TableTypeStyle 'color: red' } | Should -Throw
    }
}
