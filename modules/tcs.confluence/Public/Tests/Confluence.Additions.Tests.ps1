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

Describe 'REST additions' {
    BeforeAll {
        function Get-DecodedUri($Uri) { [System.Net.WebUtility]::UrlDecode($Uri.OriginalString) }
    }

    BeforeEach {
        Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Username 'user@contoso.com' -PersonalAccessToken 'token'
        Mock -ModuleName tcs.confluence Invoke-WebRequest {
            [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"id":"1","title":"A","name":"a"}]}' }
        }
    }

    Context 'Search-ConfluenceContent' {
        It 'Sends the CQL query to /wiki/rest/api/search and writes the results' {
            $results = @(Search-ConfluenceContent -Cql 'type = page AND label = "run book"' -Limit 50 -Expand content.space)
            $results.Count | Should -Be 1
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                (Get-DecodedUri $Uri) -eq 'https://contoso.atlassian.net/wiki/rest/api/search?cql=type = page AND label = "run book"&expand=content.space&limit=50'
            }
        }

        It 'Follows v1 next links relative to /wiki with -All' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                if ($Uri.OriginalString -like '*cursor=2*') {
                    return [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"title":"B"}],"_links":{}}' }
                }
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = '{"results":[{"title":"A"}],"_links":{"next":"/rest/api/search?cql=x&cursor=2"}}' }
            }
            (Search-ConfluenceContent -Cql 'x' -MaxQueryPages 1 -All).title | Should -Be @('A', 'B')
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/rest/api/search?cql=x&cursor=2' }
        }
    }

    Context 'Get-ConfluencePageChild' {
        It 'Gets the direct children from the v2 children endpoint' {
            $children = @(Get-ConfluencePageChild -PageId 10)
            $children.id | Should -Be @('1')
            $children[0].parentId | Should -Be '10'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/10/children?limit=250' }
        }

        It 'Returns every descendant with -Recurse' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest {
                $content = switch -Wildcard ($Uri.OriginalString) {
                    '*/pages/10/children*' { '{"results":[{"id":"11"},{"id":"12"}]}' }
                    '*/pages/11/children*' { '{"results":[{"id":"111"}]}' }
                    default { '{"results":[]}' }
                }
                [pscustomobject]@{ StatusCode = 200; StatusDescription = 'OK'; Content = $content }
            }
            $pages = @(Get-ConfluencePageChild -PageId 10 -Recurse)
            $pages.id | Should -Be @('11', '12', '111')
            ($pages | Where-Object id -EQ '111').parentId | Should -Be '11'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 4 -Exactly
        }
    }

    Context 'Labels' {
        It 'Gets labels from the v2 labels endpoint' {
            $labels = @(Get-ConfluencePageLabel -PageId 5 -Prefix global)
            $labels.name | Should -Be @('a')
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/5/labels?limit=250&prefix=global' }
        }

        It 'Adds labels with one POST of a JSON array' {
            $null = Add-ConfluencePageLabel -PageId 5 -Label 'ops'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $json = [System.Text.Encoding]::UTF8.GetString($Body)
                $Method -eq 'POST' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/rest/api/content/5/label' -and
                $json -eq '[{"prefix":"global","name":"ops"}]'
            }
        }

        It 'Rejects labels with spaces and sends nothing with -WhatIf' {
            { Add-ConfluencePageLabel -PageId 5 -Label 'two words' } | Should -Throw
            Add-ConfluencePageLabel -PageId 5 -Label ops -WhatIf
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
        }

        It 'Removes each label with DELETE ...label?name=' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 204; StatusDescription = 'No Content'; Content = '' } }
            [pscustomobject]@{ id = '5' } | Remove-ConfluencePageLabel -Label 'a b/c', 'draft' -Confirm:$false
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $Method -eq 'DELETE' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/rest/api/content/5/label?name=a+b%2Fc'
            }
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $Uri.OriginalString -like '*label?name=draft' }
        }
    }

    Context 'Attachments' {
        It 'Lists attachments from the v2 attachments endpoint' {
            $null = Get-ConfluenceAttachment -PageId 5 -FileName 'a b.pdf'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { (Get-DecodedUri $Uri) -eq 'https://contoso.atlassian.net/wiki/api/v2/pages/5/attachments?filename=a b.pdf&limit=250' }
        }

        It 'Uploads a file as multipart/form-data with the no-check header' {
            $file = Join-Path -Path $TestDrive -ChildPath 'report.txt'
            [System.IO.File]::WriteAllText($file, 'hello')
            $result = @(Add-ConfluenceAttachment -PageId 5 -Path $file -Comment 'nightly')
            $result.Count | Should -Be 1
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 1 -Exactly -ParameterFilter {
                $text = [System.Text.Encoding]::UTF8.GetString($Body)
                $Method -eq 'PUT' -and $Uri.OriginalString -eq 'https://contoso.atlassian.net/wiki/rest/api/content/5/child/attachment' -and
                $ContentType -like 'multipart/form-data; boundary=*' -and $Headers['X-Atlassian-Token'] -eq 'no-check' -and
                $text -match 'name="file"; filename="report.txt"' -and $text -match "`r`n`r`nhello`r`n" -and
                $text -match 'name="comment"' -and $text -match 'name="minorEdit"\r\n\r\ntrue'
            }
        }

        It 'Reports a missing file and uploads nothing' {
            Add-ConfluenceAttachment -PageId 5 -Path (Join-Path -Path $TestDrive -ChildPath 'missing.bin') -ErrorAction SilentlyContinue -ErrorVariable uploadError
            "$uploadError" | Should -BeLike '*File not found*'
            Should -Invoke -ModuleName tcs.confluence Invoke-WebRequest -Times 0 -Exactly
        }

        It 'Reports HTTP errors' {
            Mock -ModuleName tcs.confluence Invoke-WebRequest { [pscustomobject]@{ StatusCode = 403; StatusDescription = 'Forbidden'; Content = '{"message":"no permission"}' } }
            $file = Join-Path -Path $TestDrive -ChildPath 'x.txt'
            [System.IO.File]::WriteAllText($file, 'x')
            { Add-ConfluenceAttachment -PageId 5 -Path $file -ErrorAction Stop } | Should -Throw -ExpectedMessage '*403*no permission*'
        }
    }

    Context 'Clear-ConfluenceContext' {
        It 'Forgets the context and credential' {
            Clear-ConfluenceContext
            Get-ConfluenceContext | Should -BeNullOrEmpty
            InModuleScope tcs.confluence { $script:ConfluenceCredential } | Should -BeNullOrEmpty
            { Get-ConfluencePage -PageId 1 -ErrorAction Stop } | Should -Throw -ExpectedMessage '*Set-ConfluenceContext*'
        }

        It 'Changes nothing with -WhatIf' {
            Clear-ConfluenceContext -WhatIf
            Get-ConfluenceContext | Should -Not -BeNullOrEmpty
        }
    }
}
