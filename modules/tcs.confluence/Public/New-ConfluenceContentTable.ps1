function New-ConfluenceContentTable {
    <#
    .SYNOPSIS
        Creates an HTML table for a Confluence page from a collection of objects.

    .DESCRIPTION
        New-ConfluenceContentTable converts objects to a table: the property names of the first object
        become the header row and every object becomes a row. All objects must be of the same type and
        have the same property names.

        Cell values that are collections or objects with several properties are rendered as nested
        tables. Text that contains a web address is converted to links. Values are inserted as given,
        so they may contain markup. An empty collection returns an empty string.

    .PARAMETER TableData
        The objects to show in the table.

    .PARAMETER TableType
        The CSS class(es) of the table, for example Wrapped or "Relative Table".

    .PARAMETER TableTypeStyle
        A single inline style, for example "Width: 100%" (default), "Width: 600px" or
        "Border: 1px solid black".

    .PARAMETER NoHeader
        Do not add the header row.

    .PARAMETER HeaderStringFormatting
        Formatting for the header cells: Bold, Italic, Underline and/or Strikethrough.

    .PARAMETER HeaderAlignmentFormatting
        Text alignment of the header cells: Left (default), Center or Right.

    .PARAMETER VerticalHeader
        Render the first cell of every row as a row header (<th scope='row'>).

    .PARAMETER CellStringFormatting
        Formatting for the data cells (all but the first column): Bold, Italic, Underline and/or
        Strikethrough.

    .PARAMETER CellAlignmentFormatting
        Text alignment of the data cells: Left (default), Center or Right.

    .PARAMETER FirstCellHeaderFormat
        Wrap the first cell of each row in a heading: 0 (no heading, default) or 1 to 6.

    .PARAMETER FirstCellStringFormatting
        Formatting for the first cell of each row: Bold, Italic, Underline and/or Strikethrough.

    .PARAMETER FirstCellAlignmentFormatting
        Text alignment of the first cell of each row. Defaults to -CellAlignmentFormatting.

    .EXAMPLE
        $rows = Get-Process | Select-Object -First 5 Name, Id
        New-ConfluenceContentTable -TableData $rows -HeaderStringFormatting Bold -VerticalHeader

        Returns a table with a bold header row and the process names as row headers.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The table data to be converted.')]
        [AllowEmptyCollection()]
        [AllowNull()]
        [array]$TableData,

        [Parameter(HelpMessage = 'The type of table to be created.')]
        [ValidateSet('Wrapped', 'Wrapped Relative Table', 'Relative Table', 'Relative Table with Header', 'Wrapped Relative Table with Header', 'Wrapped Relative Table with Header and Vertical Header')]
        [string]$TableType = 'Wrapped',

        [Parameter(HelpMessage = 'The style to be applied to the table.')]
        [ValidatePattern('^(Width: \d+%|Height: \d+%|Width: \d+px|Height: \d+px|Margin: \d+px|Padding: \d+px|Border: \d+px solid [a-zA-Z]+)$')]
        [string]$TableTypeStyle = 'Width: 100%',

        [Parameter(HelpMessage = 'Do not add a header row.')]
        [switch]$NoHeader,

        [Parameter(HelpMessage = 'The text formatting to be applied to the table header.')]
        [ValidateSet('Bold', 'Italic', 'Underline', 'Strikethrough')]
        [string[]]$HeaderStringFormatting,

        [Parameter(HelpMessage = 'The text alignment to be applied to the table header.')]
        [ValidateSet('Left', 'Center', 'Right')]
        [string]$HeaderAlignmentFormatting = 'Left',

        [Parameter(HelpMessage = 'Render the first cell of each row as a row header.')]
        [switch]$VerticalHeader,

        [Parameter(HelpMessage = 'The text formatting to be applied to the table cells.')]
        [ValidateSet('Bold', 'Italic', 'Underline', 'Strikethrough')]
        [string[]]$CellStringFormatting,

        [Parameter(HelpMessage = 'The text alignment to be applied to the table cells.')]
        [ValidateSet('Left', 'Center', 'Right')]
        [string]$CellAlignmentFormatting = 'Left',

        [Parameter(HelpMessage = "The header format to be applied to the table first cell.`n0 Normal Paragraph`n1 Heading 1 (Largest)`n2 Heading 2`n3 Heading 3`n4 Heading 4`n5 Heading 5 (Smallest)`n6 Heading 6 (Quote)")]
        [ValidateSet('0', '1', '2', '3', '4', '5', '6')]
        [string]$FirstCellHeaderFormat = '0',

        [Parameter(HelpMessage = 'The text formatting to be applied to the first cell.')]
        [ValidateSet('Bold', 'Italic', 'Underline', 'Strikethrough')]
        [string[]]$FirstCellStringFormatting,

        [Parameter(HelpMessage = 'The text alignment to be applied to the first cell.')]
        [ValidateSet('Left', 'Center', 'Right')]
        [string]$FirstCellAlignmentFormatting
    )

    if ($null -eq $TableData -or $TableData.Count -eq 0) {
        return ''
    }

    $firstRow = $TableData[0]
    $firstRowType = $firstRow.GetType()
    $firstRowProperties = @($firstRow.PSObject.Properties.Name)
    foreach ($row in $TableData) {
        if ($null -eq $row -or $row.GetType() -ne $firstRowType) {
            Write-Error 'All table rows must be of the same type.'
            return
        }
        if (Compare-Object -ReferenceObject $firstRowProperties -DifferenceObject @($row.PSObject.Properties.Name) -SyncWindow 0) {
            Write-Error 'All table rows must have the same column names.'
            return
        }
    }
    if ([string]::IsNullOrEmpty($FirstCellAlignmentFormatting)) {
        $FirstCellAlignmentFormatting = $CellAlignmentFormatting
    }

    $URLFormatting = '\b((http|https):\/\/)?((www\.)?([a-zA-Z0-9-]+\.)+[a-zA-Z]{2,})(\/[a-zA-Z0-9-._~:\/?#[\]@!$&''()*+,;=]*)?\b'

    # Renders one cell value: nested table for collections/complex objects, links for URLs
    $renderValue = {
        param($Value, [bool]$DetectLinks)
        if ($null -eq $Value) { return '' }
        $isScalar = ($Value -is [string]) -or ($Value -is [ValueType])
        if (-not $isScalar) {
            if ($Value -is [System.Collections.IEnumerable] -and @($Value).Count -ge 1) {
                return (New-ConfluenceContentTable -TableData @($Value))
            }
            if (@($Value.PSObject.Properties).Count -gt 1) {
                return (New-ConfluenceContentTable -TableData @($Value))
            }
        }
        $text = $Value.ToString()
        if ($DetectLinks -and [regex]::IsMatch($text, $URLFormatting)) {
            return (New-ConfluenceContentLink -TextBlock $text)
        }
        return $text
    }

    $TableHtml = "<table class='$TableType' style='$TableTypeStyle'>"
    if (-not $NoHeader) {
        $headerOpen = Get-HtmlFormatTag -Format $HeaderStringFormatting
        $headerClose = Get-HtmlFormatTag -Format $HeaderStringFormatting -Close
        $TableHtml += '<thead><tr>'
        foreach ($header in $firstRowProperties) {
            $TableHtml += "<th style='text-align: $HeaderAlignmentFormatting;' scope='col'>$headerOpen$header$headerClose</th>"
        }
        $TableHtml += '</tr></thead>'
    }

    $firstOpen = Get-HtmlFormatTag -Format $FirstCellStringFormatting
    $firstClose = Get-HtmlFormatTag -Format $FirstCellStringFormatting -Close
    $cellOpen = Get-HtmlFormatTag -Format $CellStringFormatting
    $cellClose = Get-HtmlFormatTag -Format $CellStringFormatting -Close
    $useHeading = $FirstCellHeaderFormat -match '^[1-6]$'

    $TableHtml += '<tbody>'
    foreach ($row in $TableData) {
        $TableHtml += '<tr>'
        $isFirstCell = $true
        foreach ($cell in $row.PSObject.Properties) {
            if ($isFirstCell) {
                $cellTag = if ($VerticalHeader) { 'th' } else { 'td' }
                $scope = if ($VerticalHeader) { " scope='row'" } else { '' }
                $wrapOpen = if ($useHeading) { "<h$FirstCellHeaderFormat>" } else { '<span>' }
                $wrapClose = if ($useHeading) { "</h$FirstCellHeaderFormat>" } else { '</span>' }
                $value = & $renderValue $cell.Value $false
                $TableHtml += "<$cellTag$scope style='text-align: $FirstCellAlignmentFormatting;'>$wrapOpen$firstOpen$value$firstClose$wrapClose</$cellTag>"
                $isFirstCell = $false
            }
            else {
                $value = & $renderValue $cell.Value $true
                $TableHtml += "<td style='text-align: $CellAlignmentFormatting;'>$cellOpen$value$cellClose</td>"
            }
        }
        $TableHtml += '</tr>'
    }
    $TableHtml += '</tbody></table>'
    return $TableHtml
}
