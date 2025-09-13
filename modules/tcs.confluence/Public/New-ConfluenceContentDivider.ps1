function New-ConfluenceContentDivider {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(HelpMessage = "The type of divider to create.")]
        [ValidateSet("line", "space", "default", "dashed", "dotted", "double", "gradient")]
        [string]$Type = "line"
    )

    begin {
        $Divider = @{
            line     = "<hr>"
            space    = "<p>&nbsp;</p>"
            default  = "<hr class='default'>"
            dashed   = "<hr class='dashed'>"
            dotted   = "<hr class='dotted'>"
            double   = "<hr class='double'>"
            gradient = "<hr class='gradient'>"
        }
    }
    process {
        return $Divider[$Type]
    }
}
