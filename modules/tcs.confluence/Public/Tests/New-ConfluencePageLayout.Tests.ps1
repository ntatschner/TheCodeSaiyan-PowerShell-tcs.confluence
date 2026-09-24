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

        $result.LayoutType | Should -Be $LayoutType
        $result.ContentSections | Should -Be $Count
        $result.LayoutXml | Should -Match "^<ac:layout>"
        $result.LayoutXml | Should -Match "<ac:layout-section ac:type=`"$LayoutType`">"
        ([regex]::Matches($result.LayoutXml, '<ac:layout-cell>')).Count | Should -Be $Count
        $result.LayoutXml | Should -Match '(?s)<ac:layout-cell>\s*<p>one</p>\s*</ac:layout-cell>'
        $result.LayoutXml | Should -Match '</ac:layout>$'
    }

    It 'Only offers the section parameters the layout needs' {
        { New-ConfluencePageLayout -LayoutType single -SectionOne 'a' -SectionTwo 'b' } | Should -Throw
    }

    It 'Rejects unknown layouts' {
        { New-ConfluencePageLayout -LayoutType 'four_equal' } | Should -Throw
    }
}
