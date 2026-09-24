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

# Every builder must return well-formed storage format (XHTML), whatever text it is given.
Describe 'Builder output is well-formed XML' {
    BeforeAll {
        $hostile = "R&D <draft> `"quoted`" it's ]]> & more"
        $rows = @(
            [pscustomobject]@{ Name = $hostile; File = 'report.pdf'; Mail = 'me@contoso.com'; Link = 'https://x.com/?a=1&b=2'; Nested = [pscustomobject]@{ A = $hostile; B = 1 } }
        )
        function Test-StorageFormat([string]$Markup) {
            # Confluence storage format uses the ac: and ri: namespaces
            $document = New-Object -TypeName System.Xml.XmlDocument
            $document.LoadXml("<root xmlns:ac=`"http://atlassian.com/content`" xmlns:ri=`"http://atlassian.com/resource/identifier`">$Markup</root>")
            return $true
        }
    }

    It '<Name> output parses as XML' -ForEach @(
        @{ Name = 'New-ConfluenceContentHeader'; Build = { New-ConfluenceContentHeader -Header $hostile -Level 2 -StringFormatting Bold, Italic } }
        @{ Name = 'New-ConfluenceContentTable'; Build = { New-ConfluenceContentTable -TableData $rows -HeaderStringFormatting Bold -VerticalHeader -FirstCellHeaderFormat 3 } }
        @{ Name = 'New-ConfluenceContentTable (hashtable rows)'; Build = { New-ConfluenceContentTable -TableData @(@{ Key = $hostile }, @{ Key = 'b' }) } }
        @{ Name = 'New-ConfluenceContentTable -Raw with builder fragments'; Build = { New-ConfluenceContentTable -TableData @([pscustomobject]@{ A = (New-ConfluenceContentStatus -Text $hostile); B = (New-ConfluenceContentLink -Url 'https://x.com/?a&b' -LinkText $hostile) }) -Raw } }
        @{ Name = 'New-ConfluenceContentLink -Url'; Build = { New-ConfluenceContentLink -Url "https://x.com/?a=1&b=it's" -LinkText $hostile } }
        @{ Name = 'New-ConfluenceContentLink -TextBlock'; Build = { New-ConfluenceContentLink -TextBlock "$hostile see https://x.com/?a=1&b=2 and mailto:me@x.com." } }
        @{ Name = 'New-ConfluenceContentDivider (line)'; Build = { New-ConfluenceContentDivider } }
        @{ Name = 'New-ConfluenceContentDivider (space)'; Build = { New-ConfluenceContentDivider -Type space } }
        @{ Name = 'New-ConfluenceContentDivider (dotted)'; Build = { New-ConfluenceContentDivider -Type dotted } }
        @{ Name = 'New-ConfluenceContentInfo'; Build = { New-ConfluenceContentInfo -Title $hostile -Content $hostile -Type error } }
        @{ Name = 'New-ConfluenceContentInfo -Raw'; Build = { New-ConfluenceContentInfo -Title 't' -Content (New-ConfluenceContentCodeBlock -Content $hostile) -Raw } }
        @{ Name = 'New-ConfluenceContentCodeBlock'; Build = { New-ConfluenceContentCodeBlock -Content $hostile -Language "a<b'" } }
        @{ Name = 'New-ConfluenceContentTOC'; Build = { New-ConfluenceContentTOC -IncludeSectionNumbers } }
        @{ Name = 'New-ConfluenceContentInternalLink'; Build = { New-ConfluenceContentInternalLink -PageTitle $hostile -SpaceKey 'A&B' -HeadingLink $hostile -LinkText $hostile } }
        @{ Name = 'New-ConfluenceContentInternalLink -InternalLinkURL'; Build = { New-ConfluenceContentInternalLink -InternalLinkURL 'https://x.com/wiki/spaces/A/pages/1/R&D' -PageTitle $hostile -HeadingLink 'x' } }
        @{ Name = 'New-ConfluenceContentStatus'; Build = { New-ConfluenceContentStatus -Text $hostile -Colour Red -Subtle } }
        @{ Name = 'New-ConfluenceContentExpand'; Build = { New-ConfluenceContentExpand -Title $hostile -Content $hostile } }
        @{ Name = 'New-ConfluenceContentJiraIssue'; Build = { New-ConfluenceContentJiraIssue -JqlQuery $hostile -Columns 'key', 'a&b' -Server $hostile } }
        @{ Name = 'ConvertTo-ConfluenceHTML'; Build = { ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# $hostile`n| a | b |`n|---|---|`n| $hostile | 2 |`n- $hostile`n**$hostile**" } }
        @{ Name = 'Join-ConfluenceContent'; Build = { foreach ($separator in 'HorizontalRule', 'NewLine', 'Space', 'Tab') { Join-ConfluenceContent -ContentBlocks '<p>a</p>', '<p>b</p>' -Separator $separator } } }
        @{ Name = 'New-ConfluencePageLayout'; Build = { New-ConfluencePageLayout -Section @{ LayoutType = 'single'; Cells = (New-ConfluenceContentHeader -Header $hostile -Level 1) }, @{ LayoutType = 'three_equal'; Cells = (New-ConfluenceContentDivider), (New-ConfluenceContentInfo -Title $hostile -Content $hostile), '' } } }
    ) {
        $markup = (& $Build) -join ''
        $markup | Should -Not -BeNullOrEmpty
        { Test-StorageFormat -Markup $markup } | Should -Not -Throw
    }

    It 'Covers every exported New-ConfluenceContent* builder' {
        $tested = @('New-ConfluenceContentHeader', 'New-ConfluenceContentTable', 'New-ConfluenceContentLink', 'New-ConfluenceContentDivider',
            'New-ConfluenceContentInfo', 'New-ConfluenceContentCodeBlock', 'New-ConfluenceContentTOC', 'New-ConfluenceContentInternalLink',
            'New-ConfluenceContentStatus', 'New-ConfluenceContentExpand', 'New-ConfluenceContentJiraIssue')
        $builders = (Get-Command -Module tcs.confluence -Name 'New-ConfluenceContent*').Name
        foreach ($builder in $builders) { $tested | Should -Contain $builder }
    }
}
