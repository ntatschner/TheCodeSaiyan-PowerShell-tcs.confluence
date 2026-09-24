function New-ConfluenceContentDivider {
    <#
    .SYNOPSIS
        Creates a divider (horizontal rule or blank paragraph) for a Confluence page.

    .DESCRIPTION
        New-ConfluenceContentDivider returns an <hr> element, optionally with a style class, or an empty
        paragraph for vertical space.

    .PARAMETER Type
        The divider type: line (default), space, default, dashed, dotted, double or gradient.

    .EXAMPLE
        New-ConfluenceContentDivider -Type space

        Returns <p>&nbsp;</p>.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = 'The type of divider to create.')]
        [ValidateSet('line', 'space', 'default', 'dashed', 'dotted', 'double', 'gradient')]
        [string]$Type = 'line'
    )

    $Divider = @{
        line     = '<hr>'
        space    = '<p>&nbsp;</p>'
        default  = "<hr class='default'>"
        dashed   = "<hr class='dashed'>"
        dotted   = "<hr class='dotted'>"
        double   = "<hr class='double'>"
        gradient = "<hr class='gradient'>"
    }
    return $Divider[$Type]
}
