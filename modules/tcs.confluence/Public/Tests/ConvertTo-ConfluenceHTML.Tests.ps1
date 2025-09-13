#Requires -Modules Pester

BeforeAll {
    . "$PSScriptRoot\..\ConvertTo-ConfluenceHTML.ps1"
}

Describe 'ConvertTo-ConfluenceHTML' {
    Context 'Markdown to HTML conversion' {
        It 'should convert headings correctly' {
            $markdown = @"
# H1
## H2
"@
            $html = ConvertTo-ConfluenceHTML -InputContent $markdown -InputFormat Markdown
            $html | Should -Match '<h1>H1</h1>'
            $html | Should -Match '<h2>H2</h2>'
        }

        It 'should convert bold and italic text' {
            $markdown = "**bold** and *italic*"
            $html = ConvertTo-ConfluenceHTML -InputContent $markdown -InputFormat Markdown
            $html | Should -Be '<strong>bold</strong> and <em>italic</em>'
        }
    }
}
