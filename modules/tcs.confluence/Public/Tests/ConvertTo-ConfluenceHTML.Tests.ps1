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


Describe 'ConvertTo-ConfluenceHTML' {
    It 'Converts headings of every level' {
        $html = ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# One`n## Two`n###### Six"
        $html | Should -BeExactly "<h1>One</h1>`n<h2>Two</h2>`n<h6>Six</h6>"
    }

    It 'Converts bold and italic text' {
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent '**bold** and *italic*' |
            Should -BeExactly '<strong>bold</strong> and <em>italic</em>'
    }

    It 'Merges consecutive bullet items into one list' {
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "- a`n- b" |
            Should -BeExactly '<ul><li>a</li><li>b</li></ul>'
    }

    It 'Converts multi-line fenced code blocks, encodes them and leaves them untouched by other rules' {
        $markdown = "``````powershell`nGet-Item *.txt | Where-Object { `$_.Length -gt 0 }`n# not a heading <b>`n``````"
        $html = ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent $markdown
        $html | Should -BeExactly "<pre><code>Get-Item *.txt | Where-Object { `$_.Length -gt 0 }`n# not a heading &lt;b&gt;</code></pre>"
    }

    It 'Handles Windows line endings' {
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# Title`r`n- item" |
            Should -BeExactly "<h1>Title</h1>`n<ul><li>item</li></ul>"
    }

    It 'Converts a row into one table with one cell per column' {
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent '| a | b |' |
            Should -BeExactly '<table><tbody><tr><td>a</td><td>b</td></tr></tbody></table>'
    }

    It 'Builds one table with a header row and skips the |---| separator' {
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "| Name | Count |`n|:---|---:|`n| a | 1 |`n| **b** | 2 |" |
            Should -BeExactly '<table><tbody><tr><th>Name</th><th>Count</th></tr><tr><td>a</td><td>1</td></tr><tr><td><strong>b</strong></td><td>2</td></tr></tbody></table>'
    }

    It 'Escapes text so markup in the Markdown is shown as text' {
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# T & C`nsome <b> text" |
            Should -BeExactly "<h1>T &amp; C</h1>`nsome &lt;b&gt; text"
    }

    It 'Produces well-formed XML for a mixed document' {
        $html = ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# T & C`nsome <b> text`n| a | b |`n|---|---|`n| 1 | 2 |`n- x & y`n``````n<tag>`n``````"
        { [xml]"<root>$html</root>" } | Should -Not -Throw
    }

    It 'Rejects unsupported formats' {
        { ConvertTo-ConfluenceHTML -InputFormat 'Html' -InputContent 'x' } | Should -Throw
    }
}
