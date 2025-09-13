#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\New-ConfluencePageLayout.ps1"
}

Describe 'New-ConfluencePageLayout' {
    It 'should create a two_equal layout with two sections' {
        $section1 = '<p>Left Column</p>'
        $section2 = '<p>Right Column</p>'

        # We must use Invoke-Command to test dynamic parameters
        $result = Invoke-Command -ScriptBlock ${function:New-ConfluencePageLayout} -ArgumentList @{
            LayoutType = 'two_equal'
            SectionOne = $section1
            SectionTwo = $section2
        }

        $result.LayoutType | Should -Be 'two_equal'
        $result.ContentSections | Should -Be 2
        $result.LayoutXml | Should -Match '<ac:layout-section ac:type="two_equal">'
        $result.LayoutXml | Should -Match "<ac:layout-cell>\s*$section1\s*</ac:layout-cell>"
        $result.LayoutXml | Should -Match "<ac:layout-cell>\s*$section2\s*</ac:layout-cell>"
    }

    It 'should create a three_equal layout with three sections' {
        $result = Invoke-Command -ScriptBlock ${function:New-ConfluencePageLayout} -ArgumentList @{
            LayoutType = 'three_equal'
            SectionOne = '1'
            SectionTwo = '2'
            SectionThree = '3'
        }

        $result.ContentSections | Should -Be 3
        $result.LayoutXml | Should -Match '<ac:layout-section ac:type="three_equal">'
        $result.LayoutXml | Should -Match '<ac:layout-cell>\s*3\s*</ac:layout-cell>'
    }
}
