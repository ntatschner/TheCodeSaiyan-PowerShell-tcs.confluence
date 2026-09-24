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


Describe 'New-ConfluencePageLayout' {
    It 'Creates a <LayoutType> layout with <Count> cell(s)' -ForEach @(
        @{ LayoutType = 'single'; Count = 1 }
        @{ LayoutType = 'two_equal'; Count = 2 }
        @{ LayoutType = 'two_left_sidebar'; Count = 2 }
        @{ LayoutType = 'two_right_sidebar'; Count = 2 }
        @{ LayoutType = 'three_equal'; Count = 3 }
        @{ LayoutType = 'three_with_sidebars'; Count = 3 }
    ) {
        $sections = @{ SectionOne = '<p>one</p>'; SectionTwo = '<p>two</p>'; SectionThree = '<p>three</p>' }
        $params = @{ LayoutType = $LayoutType }
        foreach ($name in @('SectionOne', 'SectionTwo', 'SectionThree')[0..($Count - 1)]) { $params[$name] = $sections[$name] }

        $result = New-ConfluencePageLayout @params
        $result | Should -BeOfType [string]
        $result | Should -Match '^<ac:layout>'
        $result | Should -Match "<ac:layout-section ac:type=`"$LayoutType`">"
        ([regex]::Matches($result, '<ac:layout-cell>')).Count | Should -Be $Count
        $result | Should -Match '(?s)<ac:layout-cell>\s*<p>one</p>\s*</ac:layout-cell>'
        $result | Should -Match '</ac:layout>$'
        { [xml]"<root xmlns:ac='http://atlassian.com/content'>$result</root>" } | Should -Not -Throw
    }

    It 'Builds several sections in order with -Section' {
        $result = New-ConfluencePageLayout -Section @{ LayoutType = 'single'; Cells = '<h1>T</h1>' }, @{ LayoutType = 'two_equal'; Cells = '<p>a</p>', '<p>b</p>' }
        $result | Should -BeExactly '<ac:layout><ac:layout-section ac:type="single"><ac:layout-cell><h1>T</h1></ac:layout-cell></ac:layout-section><ac:layout-section ac:type="two_equal"><ac:layout-cell><p>a</p></ac:layout-cell><ac:layout-cell><p>b</p></ac:layout-cell></ac:layout-section></ac:layout>'
    }

    It 'Rejects a section with the wrong number of cells' {
        { New-ConfluencePageLayout -Section @{ LayoutType = 'three_equal'; Cells = 'a', 'b' } } | Should -Throw -ExpectedMessage '*needs 3*'
        { New-ConfluencePageLayout -LayoutType two_equal -SectionOne 'a' } | Should -Throw -ExpectedMessage '*needs 2*'
    }

    It 'Only offers the section parameters the layout needs' {
        { New-ConfluencePageLayout -LayoutType single -SectionOne 'a' -SectionTwo 'b' } | Should -Throw
    }

    It 'Rejects unknown layouts' {
        { New-ConfluencePageLayout -LayoutType 'four_equal' } | Should -Throw
    }
}
