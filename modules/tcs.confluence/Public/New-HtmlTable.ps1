#Requires -Version 3.0
function New-HtmlTable {
    <#
    .SYNOPSIS
        Creates a new HTML table or merges data into an existing HTML table string.

    .DESCRIPTION
        This advanced function generates or modifies an HTML table string.
        In 'CreateNew' mode, it builds a table from input objects with styling and optional row merging.
        In 'MergeExisting' mode, it parses an existing HTML table string, detects its columns (best effort),
        strips simple HTML tags from detected header names, applies the provided -CellStyle to ALL <td> elements
        within the existing <tbody> (overwriting existing styles), and appends new data rows (InputObject) also
        applying the -CellStyle. Potentially applies row merging relative to the last existing row.

    #>
    [CmdletBinding(DefaultParameterSetName = 'CreateNew')]
    param(
        [Parameter(ParameterSetName = 'CreateNew', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [Parameter(ParameterSetName = 'MergeExisting', Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
        [object[]]$InputObject,

        [Parameter(ParameterSetName = 'CreateNew', Position = 1)]
        [Parameter(ParameterSetName = 'MergeExisting', Position = 2)]
        [string[]]$Properties,

        [Parameter(ParameterSetName = 'CreateNew')]
        [string]$Title,

        [Parameter(ParameterSetName = 'CreateNew')]
        [string]$CssClass,

        [Parameter(ParameterSetName = 'CreateNew')]
        [hashtable]$Style,

        [Parameter(ParameterSetName = 'CreateNew')]
        [hashtable]$HeaderStyle,

        [Parameter(ParameterSetName = 'CreateNew')]
        [Parameter(ParameterSetName = 'MergeExisting')]
        [hashtable]$CellStyle,

        [Parameter(ParameterSetName = 'CreateNew')]
        [Parameter(ParameterSetName = 'MergeExisting')]
        [switch]$MergeRows,

        [Parameter(ParameterSetName = 'CreateNew')]
        [Parameter(ParameterSetName = 'MergeExisting')]
        [ValidateNotNullOrEmpty()]
        [string[]]$MergeColumns,

        [Parameter(ParameterSetName = 'MergeExisting', Mandatory = $true, Position = 1)]
        [string]$ExistingHtmlTable,

        [Parameter(ParameterSetName = 'MergeExisting', Mandatory = $true)]
        [switch]$MergeWithExisting,

        [Parameter(ParameterSetName = 'CreateNew')]
        [Parameter(ParameterSetName = 'MergeExisting')]
        [string]$NullDisplay = '',

        [Parameter(ParameterSetName = 'CreateNew')]
        [Parameter(ParameterSetName = 'MergeExisting')]
        [switch]$UseNbspForEmpty,

        [Parameter(ParameterSetName = 'CreateNew')]
        [Parameter(ParameterSetName = 'MergeExisting')]
        [switch]$PreserveHtml
    )

    Begin {
        # --- Helper Functions (moved inside Begin to satisfy advanced function structure rules) ---

        function ConvertStyleHashtableToString($StyleHashtable) {
            if ($null -eq $StyleHashtable -or $StyleHashtable.Count -eq 0) { return "" }
            $styleParts = @()
            foreach ($key in ($StyleHashtable.Keys | Sort-Object)) {
                $val = $StyleHashtable[$key] -replace '"',''
                $styleParts += "${key}: $val;"
            }
            if ($styleParts.Count -eq 0) { return "" }
            $styleString = $styleParts -join " "
            return " style=`"$styleString`""
        }

        function Get-CellValueText {
            param(
                [object]$Value,
                [switch]$PreserveHtml,
                [string]$NullDisplay,
                [switch]$UseNbspForEmpty
            )
            if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrEmpty($Value))) {
                if ($UseNbspForEmpty) { return '&nbsp;' }
                if ($NullDisplay) {
                    return $PreserveHtml ? $NullDisplay : [System.Net.WebUtility]::HtmlEncode($NullDisplay)
                }
                return ''
            }
            $stringValue = if ($Value -isnot [string]) { "$Value" } else { $Value }
            if ($PreserveHtml) { return $stringValue }
            return [System.Net.WebUtility]::HtmlEncode($stringValue)
        }

        function _NewHtmlTable_HandleCreateNew {
            param(
                [System.Collections.Generic.List[object]]$Data,
                [string[]]$Properties,
                [string]$Title,
                [string]$TableStyleAttribute,
                [string]$TableClassAttribute,
                [hashtable]$HeaderStyle,
                [hashtable]$CellStyle,
                [bool]$DoMerge,
                [string[]]$MergeColumns,
                [string]$NullDisplay,
                [bool]$UseNbspForEmpty,
                [bool]$PreserveHtml
            )
            Write-Verbose "Executing _NewHtmlTable_HandleCreateNew helper."
            $htmlBuilder = [System.Text.StringBuilder]::new()

            if ($Data.Count -eq 0) {
                Write-Warning "No input objects for new table."
                $htmlBuilder.AppendLine("<table$TableStyleAttribute$TableClassAttribute>") | Out-Null
                if ($Title) { $htmlBuilder.AppendLine("<caption>$([System.Net.WebUtility]::HtmlEncode($Title))</caption>") | Out-Null }
                if ($Properties) {
                    $htmlBuilder.AppendLine("<thead>") | Out-Null; $htmlBuilder.AppendLine("<tr>") | Out-Null
                    $headerStyleString = $(ConvertStyleHashtableToString $HeaderStyle)
                    foreach ($prop in $Properties) { $htmlBuilder.AppendLine("<th$headerStyleString>$([System.Net.WebUtility]::HtmlEncode($prop))</th>") | Out-Null }
                    $htmlBuilder.AppendLine("</tr>") | Out-Null; $htmlBuilder.AppendLine("</thead>") | Out-Null
                    $htmlBuilder.AppendLine("<tbody></tbody>") | Out-Null
                } else {
                    $htmlBuilder.AppendLine("<tbody></tbody>") | Out-Null
                }
                $htmlBuilder.AppendLine("</table>") | Out-Null
                return $htmlBuilder.ToString()
            }

            $htmlBuilder.AppendLine("<table$TableStyleAttribute$TableClassAttribute>") | Out-Null
            if ($Title) { $htmlBuilder.AppendLine("<caption>$([System.Net.WebUtility]::HtmlEncode($Title))</caption>") | Out-Null }
            $htmlBuilder.AppendLine("<thead>") | Out-Null; $htmlBuilder.AppendLine("<tr>") | Out-Null
            $headerStyleString = $(ConvertStyleHashtableToString $HeaderStyle)
            foreach ($prop in $Properties) { $htmlBuilder.AppendLine("<th$headerStyleString>$([System.Net.WebUtility]::HtmlEncode($prop))</th>") | Out-Null }
            $htmlBuilder.AppendLine("</tr>") | Out-Null; $htmlBuilder.AppendLine("</thead>") | Out-Null
            $htmlBuilder.AppendLine("<tbody>") | Out-Null
            $cellStyleString = $(ConvertStyleHashtableToString $CellStyle)
            Write-Verbose "Applying CellStyle string: '$cellStyleString'"

            $cellAttributes = @{}
            if ($DoMerge) {
                Write-Verbose "Calculating rowspans..."
                for ($rowIndex = 0; $rowIndex -lt $Data.Count; $rowIndex++) { $cellAttributes[$rowIndex] = @{} }
                foreach ($colName in $MergeColumns) {
                    $startRow = 0
                    while ($startRow -lt $Data.Count) {
                        if (-not $cellAttributes.ContainsKey($startRow)){ $cellAttributes[$startRow] = @{} }
                        if (-not $cellAttributes[$startRow].ContainsKey($colName)){ $cellAttributes[$startRow][$colName] = @{Rowspan=1;Skip=$false} }
                        $spanCount = 1
                        for ($nextRow = $startRow + 1; $nextRow -lt $Data.Count; $nextRow++) {
                            $match = $true
                            foreach($prevCol in $MergeColumns) {
                                $currentVal = $Data[$nextRow]."$prevCol"
                                $startVal = $Data[$startRow]."$prevCol"
                                if ("$currentVal" -ne "$startVal") { $match = $false; break }
                                if ($prevCol -eq $colName) { break }
                            }
                            if ($match) {
                                if(-not $cellAttributes.ContainsKey($nextRow)){$cellAttributes[$nextRow] = @{}}
                                if(-not $cellAttributes[$nextRow].ContainsKey($colName)){$cellAttributes[$nextRow][$colName] = @{Rowspan=1;Skip=$false}}
                                $spanCount++
                                $cellAttributes[$nextRow][$colName].Skip = $true
                            } else { break }
                        }
                        if ($spanCount -gt 1) { $cellAttributes[$startRow][$colName].Rowspan = $spanCount }
                        $startRow += $spanCount
                    }
                }
            }

            Write-Verbose "Generating table body rows..."
            for ($rowIndex = 0; $rowIndex -lt $Data.Count; $rowIndex++) {
                $item = $Data[$rowIndex]
                $htmlBuilder.AppendLine("<tr>") | Out-Null
                foreach ($prop in $Properties) {
                    $cellValue = $item.$prop
                    $encodedValue = Get-CellValueText -Value $cellValue -PreserveHtml:$PreserveHtml -NullDisplay $NullDisplay -UseNbspForEmpty:$UseNbspForEmpty
                    $rowspanAttr = ""
                    $skipCell = $false
                    if ($DoMerge -and $MergeColumns -contains $prop) {
                        if ($cellAttributes.ContainsKey($rowIndex) -and $cellAttributes[$rowIndex].ContainsKey($prop)) {
                            $attr = $cellAttributes[$rowIndex][$prop]
                            if ($attr.Skip) { $skipCell = $true }
                            elseif ($attr.Rowspan -gt 1) { $rowspanAttr = " rowspan=`"$($attr.Rowspan)`"" }
                        }
                    }
                    if (-not $skipCell) {
                        $htmlBuilder.AppendLine("<td$rowspanAttr$cellStyleString>$encodedValue</td>") | Out-Null
                    }
                }
                $htmlBuilder.AppendLine("</tr>") | Out-Null
            }
            $htmlBuilder.AppendLine("</tbody>") | Out-Null
            $htmlBuilder.AppendLine("</table>") | Out-Null
            Write-Verbose "Helper _NewHtmlTable_HandleCreateNew finished."
            return $htmlBuilder.ToString()
        }

        function _NewHtmlTable_HandleMergeExisting {
            param(
                [System.Collections.Generic.List[object]]$NewData,
                [string]$ExistingHtmlTable,
                [string[]]$Properties,
                [hashtable]$CellStyle,
                [bool]$DoMerge,
                [string[]]$MergeColumns,
                [hashtable]$BoundParameters,
                [string]$NullDisplay,
                [bool]$UseNbspForEmpty,
                [bool]$PreserveHtml
            )
            Write-Verbose "Executing _NewHtmlTable_HandleMergeExisting helper."
            if ($NewData.Count -eq 0) { Write-Warning "No new data provided. Returning original table."; return $ExistingHtmlTable }

            $detectedProperties = @()
            $lastExistingRowData = $null
            $tbodyStartIndex = -1; $tbodyEndIndex = -1
            $originalTbodyContent = ""
            $tbodyStartTag = "<tbody>"; $tbodyEndTag = "</tbody>"
            $cellStyleAttribute = $(ConvertStyleHashtableToString $CellStyle)
            Write-Verbose "Target CellStyle attribute for ALL tbody cells: '$cellStyleAttribute'"

            $headerMatch = [regex]::Match($ExistingHtmlTable, '(?si)<thead>\s*<tr.*?>(.*?)</tr>\s*</thead>')
            if ($headerMatch.Success) {
                Write-Verbose "Found <thead> block."
                $headerRowHtml = $headerMatch.Groups[1].Value
                $thMatches = [regex]::Matches($headerRowHtml, '(?si)<th.*?>(.*?)</th>')
                if ($thMatches.Count -gt 0) {
                    $detectedProperties = $thMatches | ForEach-Object { $_.Groups[1].Value.Trim() -replace '<.*?>','' }
                } else { Write-Warning "Found <thead> but no <th> tags inside." }
            } else { Write-Error "Failed to parse headers (<thead>) from ExistingHtmlTable. Cannot merge."; return $null }
            if ($detectedProperties.Count -eq 0) { Write-Error "No columns detected from existing header. Cannot merge."; return $null }
            Write-Verbose "Detected CLEAN Headers: $($detectedProperties -join ', ')"

            $finalProperties = $Properties
            if ($null -eq $finalProperties -or $finalProperties.Length -eq 0) {
                Write-Verbose "Using detected properties for mapping new data: $($detectedProperties -join ', ')"; $finalProperties = $detectedProperties
            } else {
                Write-Verbose "Using specified -Properties for mapping new data: $($finalProperties -join ', ')"
                if ($finalProperties.Length -ne $detectedProperties.Length) { Write-Error "Specified -Properties count ($($finalProperties.Length)) != detected columns count ($($detectedProperties.Length))."; return $null }
                if ($NewData.Count -gt 0) {
                    $firstNewItemProps = $NewData[0].PSObject.Properties.Name
                    $missingProps = $finalProperties | Where-Object { $firstNewItemProps -notcontains $_ }
                    if ($missingProps) { Write-Warning "Specified -Properties missing on first new object: $($missingProps -join ', ')." }
                }
            }

            $tbodyMatch = [regex]::Match($ExistingHtmlTable, '(?si)(<tbody.*?>)(.*?)(</tbody>)')
            if ($tbodyMatch.Success) {
                Write-Verbose "Found <tbody> block."
                $tbodyStartTag = $tbodyMatch.Groups[1].Value
                $originalTbodyContent = $tbodyMatch.Groups[2].Value
                $tbodyEndTag = $tbodyMatch.Groups[3].Value
                $tbodyStartIndex = $tbodyMatch.Groups[1].Index
                $tbodyEndIndex = $tbodyMatch.Groups[3].Index

                if ($DoMerge -and $originalTbodyContent.Trim()) {
                    Write-Verbose "Extracting last existing row data..."
                    $lastTrMatch = [regex]::Match($originalTbodyContent.Trim(), '(?si)<tr.*?>(.*?)</tr>\s*$', [System.Text.RegularExpressions.RegexOptions]::RightToLeft)
                    if ($lastTrMatch.Success) {
                        $lastRowHtml = $lastTrMatch.Groups[1].Value
                        $tdMatches = [regex]::Matches($lastRowHtml, '(?si)<td.*?>(.*?)</td>')
                        if ($tdMatches.Count -eq $detectedProperties.Length) {
                            $lastExistingRowData = [ordered]@{}
                            for ($i = 0; $i -lt $detectedProperties.Length; $i++) {
                                $colName = $detectedProperties[$i]
                                $cellValue = [System.Net.WebUtility]::HtmlDecode($tdMatches[$i].Groups[1].Value.Trim())
                                $lastExistingRowData[$colName] = $cellValue
                            }
                            Write-Verbose "Last existing row data extracted."
                        } else { Write-Warning "Could not parse correct # of <td> cells from last existing row." }
                    } else { Write-Warning "Could not find last <tr> in existing tbody." }
                }
            } else {
                Write-Error "Could not find <tbody> tags. Cannot apply styles to existing cells or reliably merge."
                return $null
            }

            $modifiedTbodyContent = $originalTbodyContent
            if ($BoundParameters.ContainsKey('CellStyle') -and $cellStyleAttribute) {
                Write-Verbose "Applying CellStyle to existing tbody content (preserving other attributes)..."
                try {
                    $modifiedTbodyContent = [Regex]::Replace(
                        $originalTbodyContent,
                        '(?si)(<td)(.*?)>(.*?)</td>',
                        {
                            param($m)
                            $existingAttributes = $m.Groups[2].Value
                            $content = $m.Groups[3].Value
                            $attributesWithoutStyle = [regex]::Replace($existingAttributes, '(?i)\s+style\s*=\s*".*?"', '', 1)
                            return "<td$attributesWithoutStyle$cellStyleAttribute>$content</td>"
                        }
                    )
                    Write-Verbose "Finished applying CellStyle to existing content."
                } catch {
                    Write-Warning "Regex error applying style to existing tbody content: $($_.Exception.Message)"
                    $modifiedTbodyContent = $originalTbodyContent
                }
            } else {
                Write-Verbose "No -CellStyle provided or style string is empty, skipping modification of existing tbody cells."
            }

            $cellAttributes = @{}
            if ($DoMerge) {
                Write-Verbose "Calculating rowspans for new data..."
                if ($null -eq $MergeColumns -or $MergeColumns.Length -eq 0) { Write-Error "Internal Error: MergeColumns null/empty."; return $null }
                $invalidMergeCols = $MergeColumns | Where-Object { $finalProperties -notcontains $_ }
                if ($invalidMergeCols) { Write-Error "MergeColumns invalid: $($invalidMergeCols -join ', ')"; return $null }
                Write-Verbose ("MergeColumns validated: {0}" -f ($MergeColumns -join ', '))
                for ($rowIndex = 0; $rowIndex -lt $NewData.Count; $rowIndex++) {
                    $cellAttributes[$rowIndex] = @{}
                    foreach ($colName in $MergeColumns) { $cellAttributes[$rowIndex][$colName] = @{ Rowspan = 1; Skip = $false } }
                }
                foreach ($colName in $MergeColumns) {
                    $startRow = 0
                    while ($startRow -lt $NewData.Count) {
                        $spanCount = 1
                        for ($nextRow = $startRow + 1; $nextRow -lt $NewData.Count; $nextRow++) {
                            $match = $true
                            foreach($prevCol in $MergeColumns) {
                                $currentVal = $NewData[$nextRow]."$prevCol"
                                $startVal = $NewData[$startRow]."$prevCol"
                                if ("$currentVal" -ne "$startVal") { $match = $false; break }
                                if ($prevCol -eq $colName) { break }
                            }
                            if ($match) { $spanCount++; $cellAttributes[$nextRow][$colName].Skip = $true } else { break }
                        }
                        if ($spanCount -gt 1) { $cellAttributes[$startRow][$colName].Rowspan = $spanCount }
                        $compareTargetRowData = if ($startRow -eq 0) { $lastExistingRowData } else { $NewData[$startRow - 1] }
                        if ($null -ne $compareTargetRowData) {
                            $matchChainWithPrevious = $true
                            foreach($prevCol in $MergeColumns) {
                                $prevCompareValue = if ($compareTargetRowData -is [hashtable]){ $compareTargetRowData[$prevCol] } else { $compareTargetRowData.$prevCol }
                                $prevCurrentValue = $NewData[$startRow]."$prevCol"
                                if ("$prevCurrentValue" -ne "$prevCompareValue") { $matchChainWithPrevious = $false; break }
                                if ($prevCol -eq $colName) { break }
                            }
                            if ($matchChainWithPrevious) {
                                Write-Verbose "  - Row Idx $startRow, Col '$colName' matches previous. Skip=True."
                                $cellAttributes[$startRow][$colName].Skip = $true
                            }
                        }
                        $startRow += $spanCount
                    }
                }
            }

            Write-Verbose "Generating HTML for new rows..."
            $newRowsHtmlBuilder = [System.Text.StringBuilder]::new()
            for ($rowIndex = 0; $rowIndex -lt $NewData.Count; $rowIndex++) {
                $item = $NewData[$rowIndex]
                $newRowsHtmlBuilder.AppendLine("<tr>") | Out-Null
                foreach ($prop in $finalProperties) {
                    $cellValue = $item.$prop
                    $encodedValue = Get-CellValueText -Value $cellValue -PreserveHtml:$PreserveHtml -NullDisplay $NullDisplay -UseNbspForEmpty:$UseNbspForEmpty
                    $rowspanAttr = ""; $skipCell = $false
                    if ($DoMerge -and $MergeColumns -contains $prop) {
                        if ($cellAttributes.ContainsKey($rowIndex) -and $cellAttributes[$rowIndex].ContainsKey($prop)) {
                            $attr = $cellAttributes[$rowIndex][$prop]
                            if ($attr.Skip) { $skipCell = $true }
                            elseif ($attr.Rowspan -gt 1) { $rowspanAttr = " rowspan=`"$($attr.Rowspan)`"" }
                        }
                    }
                    if (-not $skipCell) {
                        $newRowsHtmlBuilder.AppendLine("<td$rowspanAttr$cellStyleAttribute>$encodedValue</td>") | Out-Null
                    }
                }
                $newRowsHtmlBuilder.AppendLine("</tr>") | Out-Null
            }
            $newRowsHtmlContent = $newRowsHtmlBuilder.ToString()
            Write-Verbose "Generated new rows content."

            Write-Verbose "Reconstructing final HTML..."
            $partBeforeTbody = $ExistingHtmlTable.Substring(0, $tbodyStartIndex)
            $partAfterTbody  = $ExistingHtmlTable.Substring($tbodyEndIndex + $tbodyEndTag.Length)
            $finalHtmlBuilder = [System.Text.StringBuilder]::new()
            $finalHtmlBuilder.Append($partBeforeTbody) | Out-Null
            $finalHtmlBuilder.Append($tbodyStartTag) | Out-Null
            $finalHtmlBuilder.Append($modifiedTbodyContent) | Out-Null
            $finalHtmlBuilder.Append($newRowsHtmlContent) | Out-Null
            $finalHtmlBuilder.Append($tbodyEndTag) | Out-Null
            $finalHtmlBuilder.Append($partAfterTbody) | Out-Null
            $finalHtml = $finalHtmlBuilder.ToString()
            Write-Verbose "Helper _NewHtmlTable_HandleMergeExisting finished."
            return $finalHtml
        }

        # --- Initialization previously in Begin block ---
        Write-Verbose "[$((Get-Date).TimeOfDay)] Function Start. Parameter Set: $($PSCmdlet.ParameterSetName)"
        $newData = [System.Collections.Generic.List[object]]::new()
        $doMerge = $false

        if ($MergeRows.IsPresent -and ($null -eq $MergeColumns -or $MergeColumns.Length -eq 0)) {
            Write-Warning "MergeRows specified without MergeColumns. Merging disabled."
            $doMerge = $false
        } elseif ($MergeRows.IsPresent) {
            $doMerge = $true
        } else {
            $doMerge = $false
        }
        if (-not $doMerge -and $PSBoundParameters.ContainsKey('MergeColumns') -and $MergeColumns.Length -gt 0) {
            Write-Warning "MergeColumns specified without MergeRows. Ignoring MergeColumns."
            $MergeColumns = $null
        }

        if ($PSCmdlet.ParameterSetName -eq 'MergeExisting') {
            if ($PSBoundParameters.ContainsKey('Title')) { Write-Warning "-Title ignored with -MergeWithExisting." }
            if ($PSBoundParameters.ContainsKey('CssClass')) { Write-Warning "-CssClass ignored with -MergeWithExisting." }
            if ($PSBoundParameters.ContainsKey('Style')) { Write-Warning "-Style ignored with -MergeWithExisting." }
            if ($PSBoundParameters.ContainsKey('HeaderStyle')) { Write-Warning "-HeaderStyle ignored with -MergeWithExisting." }
        }
        Write-Verbose "Initialization complete."
    }

    Process {
        foreach ($item in $InputObject) { $newData.Add($item) }
        Write-Verbose "Processed $($newData.Count) input item(s)."
    }

    End {
        Write-Verbose "[$((Get-Date).TimeOfDay)] End block started. Dispatching to helper function..."
        if ($PSCmdlet.ParameterSetName -eq 'MergeExisting') {
            return _NewHtmlTable_HandleMergeExisting -NewData $newData `
                                                     -ExistingHtmlTable $ExistingHtmlTable `
                                                     -Properties $Properties `
                                                     -CellStyle $CellStyle `
                                                     -DoMerge $doMerge `
                                                     -MergeColumns $MergeColumns `
                                                     -BoundParameters $PSBoundParameters `
                                                     -NullDisplay $NullDisplay `
                                                     -UseNbspForEmpty:$UseNbspForEmpty `
                                                     -PreserveHtml:$PreserveHtml
        } else {
            $tableStyleAttribute = $(ConvertStyleHashtableToString $Style)
            $tableClassAttribute = ""
            if ($PSBoundParameters.ContainsKey('CssClass') -and -not [string]::IsNullOrWhiteSpace($CssClass)) {
                $safeCssClass = $CssClass.Trim() -replace '"',''
                if ($safeCssClass) { $tableClassAttribute = " class=`"$safeCssClass`"" }
            }

            $finalProperties = $Properties
            if (($null -eq $finalProperties -or $finalProperties.Length -eq 0) -and $newData.Count -gt 0) {
                $finalProperties = $newData[0].PSObject.Properties |
                    Where-Object { $_.MemberType -match '^(NoteProperty|Property|AliasProperty)$' } |
                    Select-Object -ExpandProperty Name
                Write-Verbose ("Detected properties for CreateNew: {0}" -f ($finalProperties -join ', '))
            } elseif ($newData.Count -gt 0) {
                Write-Verbose ("Using specified properties for CreateNew: {0}" -f ($finalProperties -join ', '))
            }

            if ($doMerge) {
                if ($null -eq $MergeColumns -or $MergeColumns.Length -eq 0) {
                    Write-Error "Internal Error: MergeColumns null/empty when DoMerge is true."; return $null
                }
                if ($null -eq $finalProperties) {
                    Write-Error "Cannot validate MergeColumns because properties could not be determined."; return $null
                }
                $invalidMergeCols = $MergeColumns | Where-Object { $finalProperties -notcontains $_ }
                if ($invalidMergeCols) {
                    Write-Error ("MergeColumns not valid for CreateNew: {0}. Valid properties: {1}" -f ($invalidMergeCols -join ', '), ($finalProperties -join ', '))
                    return $null
                }
                Write-Verbose ("MergeColumns validated for CreateNew: {0}" -f ($MergeColumns -join ', '))
            }

            return _NewHtmlTable_HandleCreateNew -Data $newData `
                                                 -Properties $finalProperties `
                                                 -Title $Title `
                                                 -TableStyleAttribute $tableStyleAttribute `
                                                 -TableClassAttribute $tableClassAttribute `
                                                 -HeaderStyle $HeaderStyle `
                                                 -CellStyle $CellStyle `
                                                 -DoMerge $doMerge `
                                                 -MergeColumns $MergeColumns `
                                                 -NullDisplay $NullDisplay `
                                                 -UseNbspForEmpty:$UseNbspForEmpty `
                                                 -PreserveHtml:$PreserveHtml
        }
        Write-Verbose "[$((Get-Date).TimeOfDay)] End block finished."
    } # End End block
} # End Function New-HtmlTable
