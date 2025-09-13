#Requires -Modules Pester

BeforeAll {
    # Import all content creation functions
    Get-ChildItem "$PSScriptRoot\..\New-ConfluenceContent*.ps1" | ForEach-Object { . $_.FullName }
    . "$PSScriptRoot\..\Join-ConfluenceContent.ps1"
}

Describe 'New-ConfluenceContent* and Join-ConfluenceContent Functions' {
    It 'New-ConfluenceContentHeader should create a valid h1 tag' {
        $html = New-ConfluenceContentHeader -Header 'Title' -Level 1
        $html | Should -Be '<h1>Title</h1>'
    }

    It 'New-ConfluenceContentCodeBlock should create a code block macro' {
        $html = New-ConfluenceContentCodeBlock -Content 'Write-Host "Hello"' -Language 'powershell'
        $html | Should -Match '<ac:structured-macro ac:name=.code.>'
        $html | Should -Match '<ac:parameter ac:name=.language.>powershell</ac:parameter>'
        $html | Should -Match '<!\[CDATA\[Write-Host "Hello"\]\]>'
    }

    It 'New-ConfluenceContentLink should create an anchor tag' {
        $html = New-ConfluenceContentLink -Url 'http://example.com' -LinkText 'Example'
        $html | Should -Be '<a href=''http://example.com''>Example</a>'
    }

    It 'New-ConfluenceContentTOC should create a TOC macro' {
        $html = New-ConfluenceContentTOC
        $html | Should -Match '<ac:structured-macro ac:name=.toc.>'
    }

    It 'New-ConfluenceContentDivider should create an <hr> tag' {
        $html = New-ConfluenceContentDivider -Type 'line'
        $html | Should -Be '<hr>'
    }

    It 'New-ConfluenceContentInfo should create an info panel macro' {
        $html = New-ConfluenceContentInfo -Title 'My Title' -Content 'My content' -Type 'info'
        $html | Should -Match '<div class="aui-message aui-message-info">'
        $html | Should -Match '(?s)<p class="title">.*My Title.*</p>'
        $html | Should -Match '<p>My content</p>'
    }

    It 'Join-ConfluenceContent should join blocks with a horizontal rule' {
        $content = @('<p>Block 1</p>', '<p>Block 2</p>')
        $html = Join-ConfluenceContent -ContentBlocks $content -Separator 'HorizontalRule'
        $html | Should -Be '<p>Block 1</p><hr /><p>Block 2</p>'
    }
}
