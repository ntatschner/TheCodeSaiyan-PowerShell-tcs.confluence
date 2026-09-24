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

    It 'Converts table rows' {
        ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent '| a | b |' |
            Should -BeExactly '<table><tr><td> a | b </td></tr></table>'
    }

    It 'Rejects unsupported formats' {
        { ConvertTo-ConfluenceHTML -InputFormat 'Html' -InputContent 'x' } | Should -Throw
    }
}
