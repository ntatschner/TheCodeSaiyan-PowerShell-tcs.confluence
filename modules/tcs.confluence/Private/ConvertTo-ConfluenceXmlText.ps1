function ConvertTo-ConfluenceXmlText {
    <#
    .SYNOPSIS
        Escapes text so it can be placed in Confluence storage format (XHTML) as text or an attribute.

    .DESCRIPTION
        Replaces & < > with entities, and with -Attribute also " and '. Characters that are not
        allowed in XML 1.0 (control characters other than tab, line feed and carriage return) are
        removed, so the result always parses as XML.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Text,

        [switch]$Attribute
    )

    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $clean = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', '')
    $escaped = $clean.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;')
    if ($Attribute) {
        $escaped = $escaped.Replace('"', '&quot;').Replace("'", '&#39;')
    }
    return $escaped
}

function ConvertTo-ConfluenceLinkedText {
    <#
    .SYNOPSIS
        Escapes text and turns every http, https or mailto address in it into a link.

    .DESCRIPTION
        Only addresses with a scheme (http://, https:// or mailto:) are linked, so file names such
        as report.pdf and bare e-mail addresses stay plain text. Trailing punctuation (. , ; : ! ?
        and closing brackets) is not treated as part of the address. The text between links is
        escaped with ConvertTo-ConfluenceXmlText.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $pattern = '(?i)\b(?:https?://|mailto:)[^\s<>"'']+'
    $builder = New-Object -TypeName System.Text.StringBuilder
    $position = 0
    foreach ($match in [regex]::Matches($Text, $pattern)) {
        $url = $match.Value.TrimEnd('.', ',', ';', ':', '!', '?', ')', ']', '}')
        if ($url -match '(?i)^(https?://|mailto:)$') { continue }
        $null = $builder.Append((ConvertTo-ConfluenceXmlText -Text $Text.Substring($position, $match.Index - $position)))
        $null = $builder.Append(("<a href='{0}'>{1}</a>" -f (ConvertTo-ConfluenceXmlText -Text $url -Attribute), (ConvertTo-ConfluenceXmlText -Text $url)))
        $position = $match.Index + $url.Length
    }
    $null = $builder.Append((ConvertTo-ConfluenceXmlText -Text $Text.Substring($position)))
    return $builder.ToString()
}

function ConvertTo-ConfluenceCqlString {
    <#
    .SYNOPSIS
        Returns a value as a quoted CQL string literal.

    .DESCRIPTION
        Backslashes are escaped first, then double quotes, so a value cannot end the literal early
        and change the query.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [AllowEmptyString()]
        [string]$Value
    )

    return '"' + $Value.Replace('\', '\\').Replace('"', '\"') + '"'
}
