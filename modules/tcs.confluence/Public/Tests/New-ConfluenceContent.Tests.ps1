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


Describe 'New-ConfluenceContentHeader' {
    It 'Creates a heading of the given level' {
        New-ConfluenceContentHeader -Header 'Title' -Level 1 | Should -BeExactly '<h1>Title</h1>'
    }

    It 'Nests formatting tags correctly' {
        New-ConfluenceContentHeader -Header 'T' -Level 3 -StringFormatting Bold, Italic, Underline |
            Should -BeExactly '<h3><strong><em><u>T</u></em></strong></h3>'
    }

    It 'Rejects levels outside 1-6' {
        { New-ConfluenceContentHeader -Header 'T' -Level 7 } | Should -Throw
    }
}

Describe 'New-ConfluenceContentCodeBlock' {
    It 'Creates a code macro with the given options' {
        $html = New-ConfluenceContentCodeBlock -Content 'Write-Output "Hello"' -Language 'powershell' -Theme Midnight -LineNumbers -Collapse $true
        $html | Should -Match "^<ac:structured-macro ac:name='code'>"
        $html | Should -Match "<ac:parameter ac:name='theme'>Midnight</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='linenumbers'>true</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='collapse'>true</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='language'>powershell</ac:parameter>"
        $html | Should -Match '<!\[CDATA\[Write-Output "Hello"\]\]>'
        $html | Should -Match '</ac:structured-macro>$'
    }

    It 'Defaults to no line numbers, not collapsed, language none' {
        $html = New-ConfluenceContentCodeBlock -Content 'x'
        $html | Should -Match "<ac:parameter ac:name='linenumbers'>false</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='collapse'>false</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='language'>none</ac:parameter>"
    }

    It 'Splits a CDATA terminator in the content so the macro stays well-formed' {
        $html = New-ConfluenceContentCodeBlock -Content 'a]]>b'
        $html | Should -Match ([regex]::Escape('<![CDATA[a]]]]><![CDATA[>b]]>'))
        ([regex]::Matches($html, '\]\]></ac:plain-text-body>')).Count | Should -Be 1
    }
}

Describe 'New-ConfluenceContentDivider' {
    It 'Returns <Expected> for <Type>' -ForEach @(
        @{ Type = 'line'; Expected = '<hr>' }
        @{ Type = 'space'; Expected = '<p>&nbsp;</p>' }
        @{ Type = 'dashed'; Expected = "<hr class='dashed'>" }
    ) {
        New-ConfluenceContentDivider -Type $Type | Should -BeExactly $Expected
    }

    It 'Defaults to a line' {
        New-ConfluenceContentDivider | Should -BeExactly '<hr>'
    }
}

Describe 'New-ConfluenceContentInfo' {
    It 'Creates a message block of the given type' {
        $html = New-ConfluenceContentInfo -Title 'My Title' -Content 'My content' -Type warning
        $html | Should -Match '<div class="aui-message aui-message-warning">'
        $html | Should -Match '<span class="aui-icon icon-warning"></span>'
        $html | Should -Match '(?s)<p class="title">.*My Title.*</p>'
        $html | Should -Match '<p>My content</p>'
    }

    It 'Defaults to info' {
        New-ConfluenceContentInfo -Title 't' -Content 'c' | Should -Match 'aui-message-info'
    }
}

Describe 'New-ConfluenceContentLink' {
    It 'Creates a link with text' {
        New-ConfluenceContentLink -Url 'https://example.com' -LinkText 'Example' | Should -BeExactly "<a href='https://example.com'>Example</a>"
    }

    It 'Uses the URL as the text when no text is given' {
        New-ConfluenceContentLink -Url 'https://example.com' | Should -BeExactly "<a href='https://example.com'>https://example.com</a>"
    }

    It 'Converts every URL in a text block' {
        $html = New-ConfluenceContentLink -TextBlock 'See https://example.com/a and https://contoso.com now'
        $html | Should -BeExactly "See <a href='https://example.com/a'>https://example.com/a</a> and <a href='https://contoso.com'>https://contoso.com</a> now"
    }

    It 'Leaves text without URLs unchanged' {
        New-ConfluenceContentLink -TextBlock 'no links here' | Should -BeExactly 'no links here'
    }
}

Describe 'New-ConfluenceContentTOC' {
    It 'Creates a vertical, unnumbered list by default' {
        $html = New-ConfluenceContentTOC
        $html | Should -Match "^<ac:structured-macro ac:name='toc'>"
        $html | Should -Match "<ac:parameter ac:name='type'>list</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='outline'>false</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='minLevel'>1</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='maxLevel'>6</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='style'>none</ac:parameter>"
        $html | Should -Not -Match "ac:name='include'"
    }

    It 'Maps -HorizontalList to a flat list and -IncludeSectionNumbers to outline' {
        $html = New-ConfluenceContentTOC -HorizontalList -IncludeSectionNumbers -HeadersFromLevel 2 -HeadersToLevel 3
        $html | Should -Match "<ac:parameter ac:name='type'>flat</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='outline'>true</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='minLevel'>2</ac:parameter>"
        $html | Should -Match "<ac:parameter ac:name='maxLevel'>3</ac:parameter>"
    }

    It 'Maps <Style> to the CSS list style <Css>' -ForEach @(
        @{ Style = 'Bullet'; Css = 'disc' }
        @{ Style = 'Number'; Css = 'decimal' }
        @{ Style = 'Square'; Css = 'square' }
    ) {
        New-ConfluenceContentTOC -BulletPointStyle $Style | Should -Match "<ac:parameter ac:name='style'>$Css</ac:parameter>"
    }

    It 'Omits the style for Mixed so Confluence uses its default' {
        New-ConfluenceContentTOC -BulletPointStyle Mixed | Should -Not -Match "ac:name='style'"
    }

    It 'Reports an error when the from-level is above the to-level' {
        New-ConfluenceContentTOC -HeadersFromLevel 4 -HeadersToLevel 2 -ErrorAction SilentlyContinue -ErrorVariable tocError | Should -BeNullOrEmpty
        $tocError | Should -Not -BeNullOrEmpty
    }
}

Describe 'Join-ConfluenceContent' {
    It 'Joins blocks with <Separator>' -ForEach @(
        @{ Separator = 'HorizontalRule'; Text = '<hr />' }
        @{ Separator = 'NewLine'; Text = '<br />' }
        @{ Separator = 'Space'; Text = '&nbsp;' }
        @{ Separator = 'Tab'; Text = '&emsp;' }
    ) {
        Join-ConfluenceContent -ContentBlocks '<p>1</p>', '<p>2</p>' -Separator $Separator | Should -BeExactly "<p>1</p>$Text<p>2</p>"
    }

    It 'Uses a line break by default' {
        Join-ConfluenceContent -ContentBlocks 'a', 'b' | Should -BeExactly 'a<br />b'
    }
}
