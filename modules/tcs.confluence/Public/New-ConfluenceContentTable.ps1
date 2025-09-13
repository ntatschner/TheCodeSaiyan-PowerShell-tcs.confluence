function New-ConfluenceContentTable {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = "The table data to be converted.")]
        # Removed strict non-empty validations to allow empty results (e.g. filtered collections)
        [array]$TableData,

        [Parameter(HelpMessage = "The type of table to be created.")]
        [ValidateSet("Wrapped", "Wrapped Relative Table", "Relative Table", "Relative Table with Header", "Wrapped Relative Table with Header", "Wrapped Relative Table with Header and Vertical Header")]
        # Fixed default to a value present in ValidateSet
        [string]$TableType = "Wrapped",

        [Parameter(HelpMessage = "The style to be applied to the table.")]
        [ValidateScript({ $_ -match "^(Width: \d+%|Height: \d+%|Width: \d+px|Height: \d+px|Margin: \d+px|Padding: \d+px|Border: \d+px solid [a-zA-Z]+)$" })]
        [string]$TableTypeStyle = "Width: 100%",

        [switch]$NoHeader,

        [Parameter(HelpMessage = "The text formatting to be applied to the table header.")]
        [ValidateSet("Bold", "Italic", "Underline", "Strikethrough")]
        [string[]]$HeaderStringFormatting,

        [Parameter(HelpMessage = "The text formatting to be applied to the table header.")]
        [ValidateSet("Left", "Center", "Right")]
        [string]$HeaderAlignmentFormatting = "Left",

        [switch]$VerticalHeader,

        [Parameter(HelpMessage = "The text formatting to be applied to the table cells.")]
        [ValidateSet("Bold", "Italic", "Underline", "Strikethrough")]
        [string[]]$CellStringFormatting,

        [Parameter(HelpMessage = "The text formatting to be applied to the table cells.")]
        [ValidateSet("Left", "Center", "Right")]
        [string]$CellAlignmentFormatting = "Left",

        [Parameter(HelpMessage = "The header format to be applied to the table first cell.`n0 Normal Paragraph`n1 Heading 1 (Largest)`n2 Heading 2`n3 Heading 3`n4 Heading 4`n5 Heading 5 (Smallest)`n6 Heading 6 (Quote)")]
        [ValidateSet("0", "1", "2", "3", "4", "5", "6")]
        $FirstCellHeaderFormat = 0,

        [Parameter(HelpMessage = "The text formatting to be applied to the first cell.")]
        [ValidateSet("Bold", "Italic", "Underline", "Strikethrough")]
        [string[]]$FirstCellStringFormatting,

        [Parameter(HelpMessage = "The text alignment to be applied to the first cell.")]
        [ValidateSet("Left", "Center", "Right")]
        [string]$FirstCellAlignmentFormatting
    )
    
    begin {
        # Gracefully handle null / empty
        if (-not $TableData -or $TableData.Count -eq 0) {
            return ""
        }
        $firstRow = $TableData[0]
        $firstRowType = $firstRow.GetType()
        $firstRowCount = $firstRow.Count
        $firstRowProperties = $firstRow.PSObject.Properties.Name

        foreach ($row in $TableData) {
            if ($row.GetType() -ne $firstRowType) {
                Write-Error "All table rows must be of the same type."
                return
            }
            if ($row.Count -ne $firstRowCount) {
                Write-Error "All table rows must have the same number of columns."
                return
            }
            if (Compare-Object -DifferenceObject $row.PSObject.Properties.Name -ReferenceObject $firstRowProperties -ErrorAction SilentlyContinue) {
                Write-Error "All table rows must have the same column names."
                return
            }
        }
        $URLFormatting = '\b((http|https):\/\/)?((www\.)?([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})(\/[a-zA-Z0-9-._~:\/?#[\]@!$&''()*+,;=]*)?\b'
    }
    process {
        # Quote class attribute (TableType may contain spaces = multiple classes)
        $TableHtml = "<table class='$TableType' style='$TableTypeStyle'>"
        if (-not $NoHeader) {
            $TableHtml += "<thead><tr>"
            foreach ($header in $firstRowProperties) {
                # Fixed malformed attribute (scope was inside style); scope='col'
                $TableHtml += "<th style='text-align: $HeaderAlignmentFormatting;' scope='col'>"
                if ($HeaderStringFormatting) {
                    foreach ($format in $HeaderStringFormatting) {
                        switch ($format) {
                            "Bold" {
                                $TableHtml += "<strong>"
                            }
                            "Italic" {
                                $TableHtml += "<em>"
                            }
                            "Underline" {
                                $TableHtml += "<u>"
                            }
                            "Strikethrough" {
                                $TableHtml += "<s>"
                            }
                        }
                    }
                }
                $TableHtml += $header
                if ($HeaderStringFormatting) {
                    foreach ($format in $HeaderStringFormatting) {
                        switch ($format) {
                            "Bold" {
                                $TableHtml += "</strong>"
                            }
                            "Italic" {
                                $TableHtml += "</em>"
                            }
                            "Underline" {
                                $TableHtml += "</u>"
                            }
                            "Strikethrough" {
                                $TableHtml += "</s>"
                            }
                        }
                    }
                }
                $TableHtml += "</th>"
            }
            $TableHtml += "</tr></thead>"
        }
        $TableHtml += "<tbody>"
        foreach ($row in $TableData) {
            $TableHtml += "<tr>"
            $isFirstCell = $true
            foreach ($cell in $row.PSObject.Properties) {
                if ($isFirstCell) {
                    if ([string]::IsNullOrEmpty($FirstCellAlignmentFormatting)) {
                        $FirstCellAlignmentFormatting = $CellAlignmentFormatting
                    }
                    # Support vertical header rows: use <th scope='row'> for first column when -VerticalHeader
                    if ($VerticalHeader) {
                        $TableHtml += "<th scope='row' style='text-align: $FirstCellAlignmentFormatting;'>"
                    } else {
                        $TableHtml += "<td style='text-align: $FirstCellAlignmentFormatting;'>"
                    }
                    # Avoid invalid <h0>; only wrap if 1..6 else use a span wrapper
                    $useHeading = ($FirstCellHeaderFormat -match '^[1-6]$')
                    if ($useHeading) {
                        $TableHtml += "<h$FirstCellHeaderFormat>"
                    }
                    else {
                        $TableHtml += "<span>"
                    }
                    if ($FirstCellStringFormatting) {
                        foreach ($format in $FirstCellStringFormatting) {
                            switch ($format) {
                                "Bold" {
                                    $TableHtml += "<strong>"
                                }
                                "Italic" {
                                    $TableHtml += "<em>"
                                }
                                "Underline" {
                                    $TableHtml += "<u>"
                                }
                                "Strikethrough" {
                                    $TableHtml += "<s>"
                                }
                            }
                        }
                    }
                    if ($null -ne $cell.Value) {
                        if ($cell.Value.GetType().Name -ne "String") {
                            if ($cell.Value -is [System.Collections.IEnumerable] -and $cell.Value.Count -ge 1) {
                                $TableHtml += $(New-ConfluenceContentTable -TableData $cell.Value)
                            }
                            elseif ($cell.Value.PSObject.Properties.Count -gt 1) {
                                # Handle object with multiple properties
                                $TableHtml += $(New-ConfluenceContentTable -TableData $cell.Value)
                            }
                        } else {
                            $TableHtml += $cell.Value
                        }
                    }
                    else {
                        $TableHtml += ""
                    }
        
                    # Close formatting tags already handled later; now close heading/span
                    if ($useHeading) {
                        $TableHtml += "</h$FirstCellHeaderFormat>"
                    }
                    else {
                        $TableHtml += "</span>"
                    }
                    if ($VerticalHeader) {
                        $TableHtml += "</th>"
                    } else {
                        $TableHtml += "</td>"
                    }
                    $isFirstCell = $false
                }
                else {
                    $TableHtml += "<td style='text-align: $CellAlignmentFormatting;'>"
        
                    if ($CellStringFormatting) {
                        foreach ($format in $CellStringFormatting) {
                            switch ($format) {
                                "Bold" {
                                    $TableHtml += "<strong>"
                                }
                                "Italic" {
                                    $TableHtml += "<em>"
                                }
                                "Underline" {
                                    $TableHtml += "<u>"
                                }
                                "Strikethrough" {
                                    $TableHtml += "<s>"
                                }
                            }
                        }
                    }
                    if ($null -ne $cell.Value) {
                        if ($cell.Value.GetType().Name -ne "String") {
                            if ($cell.Value -is [System.Collections.IEnumerable] -and $cell.Value.Count -ge 1) {
                                $TableHtml += $(New-ConfluenceContentTable -TableData $cell.Value)
                            }
                            elseif ($cell.Value.PSObject.Properties.Count -gt 1) {
                                # Handle object with multiple properties
                                $TableHtml += $(New-ConfluenceContentTable -TableData $cell.Value)
                            }
                            else {
                                if ([regex]::IsMatch($cell.Value.ToString(), $URLFormatting)) {
                                    $TableHtml += $(New-ConfluenceContentLink -TextBlock $cell.Value.ToString())
                                }
                                else {
                                    $TableHtml += $cell.Value.ToString()
                                }
                            }
                        }
                        else {
                            if ([regex]::IsMatch($cell.Value, $URLFormatting)) {
                                $TableHtml += $(New-ConfluenceContentLink -TextBlock $cell.Value)
                            }
                            else {
                                $TableHtml += $cell.Value
                            }
                        }
                    }
                    else {
                        $TableHtml += ""
                    }
    
        
                    if ($CellStringFormatting) {
                        foreach ($format in [System.Linq.Enumerable]::Reverse($CellStringFormatting)) {
                            switch ($format) {
                                "Bold" {
                                    $TableHtml += "</strong>"
                                }
                                "Italic" {
                                    $TableHtml += "</em>"
                                }
                                "Underline" {
                                    $TableHtml += "</u>"
                                }
                                "Strikethrough" {
                                    $TableHtml += "</s>"
                                }
                            }
                        }
                    }
        
                    $TableHtml += "</td>"
                }
            }
            # Added missing row closing tag
            $TableHtml += "</tr>"
        }
        $TableHtml += "</tbody>"
        # Always close table (was conditional previously)
        $TableHtml += "</table>"
        return $TableHtml
    }
}
