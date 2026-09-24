function New-ConfluenceContentTable {
    <#
    .SYNOPSIS
        Creates an HTML table for a Confluence page from a collection of objects.

    .DESCRIPTION
        New-ConfluenceContentTable converts objects to a table: the property names of the first object
        become the header row and every object becomes a row. All objects must be of the same type and
        have the same property names. Hashtables and ordered dictionaries are also accepted as rows;
        their keys are the columns.

        Cell values that are collections, dictionaries or objects with several properties are rendered
        as nested tables, up to three levels deep (deeper values are shown as text). In the data cells (all but the first column) addresses with a scheme
        (http://, https:// or mailto:) are converted to links; file names such as report.pdf and bare
        e-mail addresses are not. Header names and cell values are escaped, so the result is always
        well-formed storage format; use -Raw to insert cell values that already are storage-format
        markup. An empty collection returns an empty string.

    .PARAMETER TableData
        The objects (or hashtables / ordered dictionaries) to show in the table.

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

    .PARAMETER Raw
        Insert cell values as given, without escaping them or converting addresses to links. Use it
        when the values are storage-format fragments, for example links or status lozenges built with
        the other New-ConfluenceContent* functions. Header names are still escaped.

    .EXAMPLE
        New-ConfluenceContentTable -TableData @([ordered]@{ Service = 'web'; Docs = 'https://contoso.com/web' })

        Returns a table from an ordered dictionary; the address becomes a link.

    .EXAMPLE
        New-ConfluenceContentTable -TableData @([pscustomobject]@{ Name = 'web'; State = (New-ConfluenceContentStatus -Text 'OK' -Colour Green) }) -Raw

        Returns a table whose State cells contain status lozenges.

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
        [string]$FirstCellAlignmentFormatting,

        [Parameter(HelpMessage = 'Insert cell values as storage-format markup without escaping them.')]
        [switch]$Raw
    )

    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        if ($null -eq $TableData -or $TableData.Count -eq 0) {
            return ''
        }

        # Rows are objects (columns = property names) or dictionaries (columns = keys)
        $getColumns = {
            param($Row)
            if ($Row -is [System.Collections.IDictionary]) { return ($Row.Keys | ForEach-Object { [string]$_ }) }
            return $Row.PSObject.Properties.Name
        }
        $getValue = {
            param($Row, [string]$Column)
            if ($Row -is [System.Collections.IDictionary]) { return , $Row[$Column] }
            return , $Row.PSObject.Properties[$Column].Value
        }

        $firstRow = $TableData[0]
        if ($null -eq $firstRow) {
            Write-Error 'Table rows must not be null.'
            return
        }
        $isDictionary = $firstRow -is [System.Collections.IDictionary]
        $firstRowType = $firstRow.GetType()
        $firstRowProperties = @(& $getColumns $firstRow)
        foreach ($row in $TableData) {
            if ($null -eq $row) {
                Write-Error 'Table rows must not be null.'
                return
            }
            if ($isDictionary) {
                if ($row -isnot [System.Collections.IDictionary]) {
                    Write-Error 'All table rows must be of the same type.'
                    return
                }
                $rowColumns = @(& $getColumns $row)
                if ($rowColumns.Count -ne $firstRowProperties.Count -or @($rowColumns | Where-Object { $firstRowProperties -notcontains $_ }).Count -gt 0) {
                    Write-Error 'All table rows must have the same column names.'
                    return
                }
                continue
            }
            if ($row.GetType() -ne $firstRowType) {
                Write-Error 'All table rows must be of the same type.'
                return
            }
            if (Compare-Object -ReferenceObject $firstRowProperties -DifferenceObject @(& $getColumns $row) -SyncWindow 0) {
                Write-Error 'All table rows must have the same column names.'
                return
            }
        }
        if ([string]::IsNullOrEmpty($FirstCellAlignmentFormatting)) {
            $FirstCellAlignmentFormatting = $CellAlignmentFormatting
        }

        # Renders one cell value: nested table for collections/complex objects, escaped text otherwise
        $renderValue = {
            param($Value, [bool]$DetectLinks)
            if ($null -eq $Value) { return '' }
            $isScalar = ($Value -is [string]) -or ($Value -is [ValueType])
            # Nested tables stop at three levels, so self-referencing objects cannot recurse forever
            if (-not $isScalar -and $script:ConfluenceTableNesting -lt 3) {
                $nestedRows = $null
                if ($Value -is [System.Collections.IDictionary]) {
                    $nestedRows = @(, $Value)
                }
                elseif ($Value -is [System.Collections.IEnumerable] -and @($Value).Count -ge 1) {
                    $nestedRows = @($Value)
                }
                elseif (@($Value.PSObject.Properties).Count -gt 1) {
                    $nestedRows = @($Value)
                }
                if ($null -ne $nestedRows) {
                    $script:ConfluenceTableNesting++
                    try {
                        return (New-ConfluenceContentTable -TableData $nestedRows -Raw:$Raw)
                    }
                    finally {
                        $script:ConfluenceTableNesting--
                    }
                }
            }
            $text = $Value.ToString()
            if ($Raw) { return $text }
            if ($DetectLinks) { return (ConvertTo-ConfluenceLinkedText -Text $text) }
            return (ConvertTo-ConfluenceXmlText -Text $text)
        }

        $TableHtml = "<table class='$TableType' style='$TableTypeStyle'>"
        if (-not $NoHeader) {
            $headerOpen = Get-HtmlFormatTag -Format $HeaderStringFormatting
            $headerClose = Get-HtmlFormatTag -Format $HeaderStringFormatting -Close
            $TableHtml += '<thead><tr>'
            foreach ($header in $firstRowProperties) {
                $TableHtml += "<th style='text-align: $HeaderAlignmentFormatting;' scope='col'>$headerOpen$(ConvertTo-ConfluenceXmlText -Text $header)$headerClose</th>"
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
            foreach ($column in $firstRowProperties) {
                $cellValue = & $getValue $row $column
                if ($isFirstCell) {
                    $cellTag = if ($VerticalHeader) { 'th' } else { 'td' }
                    $scope = if ($VerticalHeader) { " scope='row'" } else { '' }
                    $wrapOpen = if ($useHeading) { "<h$FirstCellHeaderFormat>" } else { '<span>' }
                    $wrapClose = if ($useHeading) { "</h$FirstCellHeaderFormat>" } else { '</span>' }
                    $value = & $renderValue $cellValue $false
                    $TableHtml += "<$cellTag$scope style='text-align: $FirstCellAlignmentFormatting;'>$wrapOpen$firstOpen$value$firstClose$wrapClose</$cellTag>"
                    $isFirstCell = $false
                }
                else {
                    $value = & $renderValue $cellValue $true
                    $TableHtml += "<td style='text-align: $CellAlignmentFormatting;'>$cellOpen$value$cellClose</td>"
                }
            }
            $TableHtml += '</tr>'
        }
        $TableHtml += '</tbody></table>'
        return $TableHtml
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        throw
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
