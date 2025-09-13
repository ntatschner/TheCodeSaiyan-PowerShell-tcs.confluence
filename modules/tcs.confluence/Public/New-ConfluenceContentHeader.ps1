function New-ConfluenceContentHeader {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "String to be used as the header.")]
        [string]$Header,

        [Parameter(Mandatory = $true, HelpMessage = "The level of the header. Valid values: 1-6.")]
        [ValidateRange(1, 6)]
        [int]$Level,

        [Parameter(HelpMessage = "The text formatting to be applied to the header.")]
        [ValidateSet("Bold", "Italic", "Underline", "Strikethrough")]
        [string[]]$StringFormatting
    )

    process {
        $HeaderTagStart = "<h$Level>"
        $HeaderTagEnd = "</h$Level>"

        if ($StringFormatting) {
            foreach ($format in $StringFormatting) {
                switch ($format) {
                    "Bold" {
                        $HeaderTagStart += "<strong>"
                        $HeaderTagEnd = "</strong>" + $HeaderTagEnd
                    }
                    "Italic" {
                        $HeaderTagStart += "<em>"
                        $HeaderTagEnd = "</em>" + $HeaderTagEnd
                    }
                    "Underline" {
                        $HeaderTagStart += "<u>"
                        $HeaderTagEnd = "</u>" + $HeaderTagEnd
                    }
                    "Strikethrough" {
                        $HeaderTagStart += "<s>"
                        $HeaderTagEnd = "</s>" + $HeaderTagEnd
                    }
                }
            }
        }

        $HeaderHtml = "$HeaderTagStart$Header$HeaderTagEnd"
        return $HeaderHtml
    }
}
