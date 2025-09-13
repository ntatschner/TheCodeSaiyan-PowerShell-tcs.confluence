#Requires -Version 3.0
function Get-HtmlTableRowData {
    <#
    .SYNOPSIS
        Extracts row data from an HTML table string into PowerShell objects, attempting to parse common Confluence macros.

    .DESCRIPTION
        This function parses an HTML string containing one or more tables and attempts
        to extract the data from the cells (<td>) within the body (<tbody>) of a specified table.
        It uses regular expressions for parsing, which has limitations with complex or
        malformed HTML.

        It attempts to detect column headers (<th> in <thead>) and optionally row headers
        (first <th> in a <tbody> row).

        Crucially, it includes logic to specifically parse common Confluence macros found within
        cell content before general tag stripping:
        - User Links (<ac:userlink>): Replaces with the user's display name.
        - Confluence Links (<ac:link>): Replaces with the link's display text.
        - Standard Links (<a>): Replaces with the link's display text.
        - Jira Issues (<ac:structured-macro name="jira">): Replaces with the Jira issue key.

    .PARAMETER HtmlContent
        A string containing the HTML source code that includes the table(s). Often from a Confluence export or API.

    .PARAMETER TableIndex
        The zero-based index of the table to extract data from if multiple tables exist
        in the HtmlContent. Defaults to 0 (the first table found).

    .PARAMETER NoHeader
        Switch parameter. If specified, the function will not attempt to read column headers
        from <thead><th> tags. It will use default property names like "Column1", "Column2", etc.,
        for the data cells (<td>). Row header detection still occurs unless -NoRowHeader is also specified.

    .PARAMETER NoRowHeader
        Switch parameter. If specified, the function will not treat the first <th> element in a
        <tbody> row as a special row header.

    .PARAMETER RowHeaderColumnName
        The property name to use in the output objects for the data extracted from row headers
        (first <th> cell in a <tbody> row). Defaults to "RowHeader". Must be a valid simple variable name.

    .PARAMETER DecodeHtmlEntities
        Switch parameter. If specified, attempts to decode HTML entities (like &, <,  )
        found within header and table cell data using [System.Web.HttpUtility]::HtmlDecode.
        Defaults to $true (decoding enabled). Specify -DecodeHtmlEntities:$false to disable.

    .EXAMPLE
        # Example 1: Table with Confluence Macros
        $confluenceHtml = @"
        <table>
         <thead><tr><th>Task</th><th>Assignee</th><th>Status</th><th>Related Link</th></tr></thead>
         <tbody>
           <tr>
             <td><ac:structured-macro ac:name="jira" ac:schema-version="1" ac:macro-id="..."><ac:parameter ac:name="key">PROJ-123</ac:parameter></ac:structured-macro></td>
             <td><ac:userlink ac:username="jdoe" ac:userkey="...">John Doe</ac:userlink></td>
             <td>Done</td>
             <td><ac:link><ri:page ri:content-title="Documentation"/><ac:plain-text-link-body><![CDATA[Project Docs]]></ac:plain-text-link-body></ac:link></td>
           </tr>
           <tr>
             <td><ac:structured-macro ac:name="jira"><ac:parameter ac:name="key">PROJ-456</ac:parameter></ac:structured-macro></td>
             <td><ac:userlink ac:username="asmith">Alice Smith</ac:userlink></td>
             <td>In Progress</td>
             <td><a href='http://example.com'>External Site</a></td>
           </tr>
         </tbody>
        </table>
        "@
        Get-HtmlTableRowData -HtmlContent $confluenceHtml

        # Expected Output (may vary slightly based on exact macro rendering):
        # Task     Assignee    Status      Related Link
        # ----     --------    ------      ------------
        # PROJ-123 John Doe    Done        Project Docs
        # PROJ-456 Alice Smith In Progress External Site

    .EXAMPLE
        # Example 2: No column headers, but row headers and macros
        $htmlRowMacro = @"
        <table>
            <tr><th>PROJ-123</th><td><ac:userlink>John Doe</ac:userlink></td><td>Done</td></tr>
            <tr><th>PROJ-456</th><td><ac:userlink>Alice Smith</ac:userlink></td><td>WIP</td></tr>
        </table>
        "@
        Get-HtmlTableRowData -HtmlContent $htmlRowMacro -NoHeader -RowHeaderColumnName "JiraKey"

        # Output:
        # JiraKey  Column1     Column2
        # -------  -------     -------
        # PROJ-123 John Doe    Done
        # PROJ-456 Alice Smith WIP


    .OUTPUTS
        PSCustomObject[] - An array of PowerShell custom objects, where each object represents a row.
                         Returns $null if parsing fails at a critical step. Returns @() if no data rows found.

    .NOTES
        Author: AI Assistant
        Date: 2023-10-27
        - WARNING: Regex-based parsing of Confluence Storage Format / HTML is VERY fragile. It relies on specific
          tag structures observed in common exports. Changes in Confluence versions or complex macro usage WILL break this.
          Consider dedicated Confluence API clients or libraries that handle storage format conversion if robustness is critical.
        - Macro parsing attempts to extract the most common user-visible text (display names, link text, Jira keys).
        - Assumes standard table structure. Row headers are assumed to be the first <th> in a <tbody> row.
        - Basic HTML tags *remaining after macro processing* are stripped.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$HtmlContent,

        [Parameter()]
        [int]$TableIndex = 0,

        [Parameter()]
        [switch]$NoHeader,

        [Parameter()]
        [switch]$NoRowHeader,

        [Parameter()]
        [string]$RowHeaderColumnName = "RowHeader",

        [Parameter()]
        [Alias('Decode')]
        [bool]$DecodeHtmlEntities = $true
    )

    Begin {
        Write-Verbose "[$((Get-Date).TimeOfDay)] Function Start."
        # Validate RowHeaderColumnName using Regex
        if ($RowHeaderColumnName -notmatch '^[a-zA-Z_][a-zA-Z0-9_]*$') {
            Write-Error "Invalid -RowHeaderColumnName specified: '$RowHeaderColumnName'."; return $null
        }
        try { Add-Type -AssemblyName System.Web -ErrorAction Stop } catch { Write-Warning "System.Web assembly load failed. HTML entity decoding might fail." }
    }

    Process {
        # Process block is empty
    }

    End {
        Write-Verbose "[$((Get-Date).TimeOfDay)] End block started. Processing TableIndex $TableIndex."

        # 1. Find the specified table
        $tableMatches = [regex]::Matches($HtmlContent, '(?si)<table.*?>.*?</table>')
        if ($tableMatches.Count -eq 0) { Write-Warning "No <table> elements found."; return $null }
        if ($TableIndex -ge $tableMatches.Count) { Write-Warning "TableIndex $TableIndex out of bounds (Found $($tableMatches.Count))."; return $null }
        $targetTableHtml = $tableMatches[$TableIndex].Value
        Write-Verbose "Found target table (Index $TableIndex)."

        # 2. Extract Column Headers (if not -NoHeader)
        [string[]]$columnHeaders = @()
        $maxDataCellsFound = 0

        if (-not $NoHeader.IsPresent) {
            Write-Verbose "Attempting to detect column headers from <thead>..."
            $theadMatch = [regex]::Match($targetTableHtml, '(?si)<thead>(.*?)</thead>')
            if ($theadMatch.Success) {
                $headerRowMatch = [regex]::Match($theadMatch.Groups[1].Value, '(?si)<tr.*?>(.*?)</tr>')
                if ($headerRowMatch.Success) {
                    $thMatches = [regex]::Matches($headerRowMatch.Groups[1].Value, '(?si)<th.*?>(.*?)</th>')
                    if ($thMatches.Count -gt 0) {
                         $startIndex = 0; if ($thMatches.Count -gt 1 -and [string]::IsNullOrWhiteSpace($thMatches[0].Groups[1].Value -replace '<.*?>','')) { Write-Verbose "Skipping first potential empty corner header cell."; $startIndex = 1 }
                         # Process headers: Decode entities FIRST, then strip tags
                         $columnHeaders = $thMatches[$startIndex..($thMatches.Count - 1)] | ForEach-Object {
                            $headerText = $_.Groups[1].Value.Trim()
                            if ($DecodeHtmlEntities) { try { $headerText = [System.Web.HttpUtility]::HtmlDecode($headerText) } catch { Write-Warning "Header decode error: $($_.Exception.Message)" } }
                            $headerText -replace '<.*?>','' # Strip tags last
                         }
                         Write-Verbose "Detected Column Headers: $($columnHeaders -join ', ')"
                    } else { Write-Verbose "Found <thead>/<tr> but no <th> tags. Using default names." }
                } else { Write-Verbose "Found <thead> but no <tr> tags. Using default names." }
            } else { Write-Verbose "No <thead> found. Using default names." }
        } else { Write-Verbose "-NoHeader specified. Using default column names for data cells." }

        # 3. Extract Body Content
        Write-Verbose "Extracting content from <tbody>..."
        $tbodyMatch = [regex]::Match($targetTableHtml, '(?si)<tbody.*?>(.*?)</tbody>')
        $tbodyContent = ""
        if (-not $tbodyMatch.Success) {
             Write-Warning "No <tbody> found. Attempting fallback extraction (may be inaccurate)."
             # Fallback logic (remains simplified)
             $headerEndIndex = 0; $headerRowMatchLocal = $null # Use local var to avoid conflict
             if ($theadMatch.Success) { $headerEndIndex = $theadMatch.Index + $theadMatch.Length }
             elseif($headerRowMatchLocal = [regex]::Match($targetTableHtml, '(?si)<tr.*?>(.*?)</tr>')) { if([regex]::IsMatch($headerRowMatchLocal.Groups[1].Value, '(?si)<th.*?>')){ $headerEndIndex = $headerRowMatchLocal.Index + $headerRowMatchLocal.Length } }
             else { $tableTagMatch = [regex]::Match($targetTableHtml, '(?si)<table.*?>'); if ($tableTagMatch.Success) { $headerEndIndex = $tableTagMatch.Length } }
             $endTableMatch = [regex]::Match($targetTableHtml, '(?si)</table>\s*$', [System.Text.RegularExpressions.RegexOptions]::RightToLeft)
             if ($endTableMatch.Success -and $headerEndIndex -lt $endTableMatch.Index) { $tbodyContent = $targetTableHtml.Substring($headerEndIndex, $endTableMatch.Index - $headerEndIndex) }
             else { Write-Warning "Could not determine body content without <tbody>."; $tbodyContent = "" }
        } else { $tbodyContent = $tbodyMatch.Groups[1].Value; Write-Verbose "Found <tbody> content." }


        # 4. Extract Rows
        $outputObjects = [System.Collections.Generic.List[PSCustomObject]]::new()
        $rowMatches = [regex]::Matches($tbodyContent, '(?si)<tr.*?>(.*?)</tr>')
        if ($rowMatches.Count -eq 0) { Write-Warning "No data rows (<tr>) found within table body."; return @() }
        Write-Verbose "Found $($rowMatches.Count) data row(s) (<tr>)."


        # --- Pre-calculate max DATA cells (<td>) if default COL headers are needed ---
        if ($columnHeaders.Count -eq 0) {
             Write-Verbose "Calculating maximum DATA cell (<td>) count for default headers..."
             foreach ($rowMatch in $rowMatches) { $dataCellMatches = [regex]::Matches($rowMatch.Groups[1].Value, '(?si)<td.*?>(.*?)</td>'); if ($dataCellMatches.Count -gt $maxDataCellsFound) { $maxDataCellsFound = $dataCellMatches.Count } }
             Write-Verbose "Maximum data cells (<td>) found in a row: $maxDataCellsFound"
             if ($maxDataCellsFound -eq 0 -and $rowMatches.Count -gt 0) { Write-Warning "Found rows but no data cells (<td>) within them." }
        }
        # --- Generate default COL headers if needed ---
        if ($columnHeaders.Count -eq 0 -and $maxDataCellsFound -gt 0) { $columnHeaders = 1..$maxDataCellsFound | ForEach-Object { "Column$_" }; Write-Verbose "Generated default column headers: $($columnHeaders -join ', ')" }


        # --- Process each row ---
        $rowCount = 0
        foreach ($rowMatch in $rowMatches) {
            $rowCount++
            $rowHtml = $rowMatch.Groups[1].Value.Trim() # Inner HTML of the <tr>
            $rowData = [ordered]@{}
            $rowHeaderValue = $null
            $dataCellsHtml = $rowHtml # Assume all cells are data initially

            # --- Process Row Header Cell (if present and not disabled) ---
            if (-not $NoRowHeader.IsPresent) {
                $firstCellMatch = [regex]::Match($rowHtml, '(?si)^\s*<th(\s+[^>]*?)?>(.*?)</th>')
                if ($firstCellMatch.Success) {
                    Write-Verbose "Row $rowCount`: Found row header (<th>)."
                    # Process header content: Decode FIRST, then handle macros, then strip remaining tags
                    $rawHeaderContent = $firstCellMatch.Groups[2].Value
                    $processedHeaderContent = $rawHeaderContent # Start with raw

                    if ($DecodeHtmlEntities) { try { $processedHeaderContent = [System.Web.HttpUtility]::HtmlDecode($processedHeaderContent) } catch { Write-Warning "Row $rowCount`: Error decoding row header: $($_.Exception.Message)." } }

                    $processedHeaderContent = $processedHeaderContent -replace '(?si)<ac:userlink.*?>\s*(.*?)\s*</ac:userlink>', '$1' `
                                                                   -replace '(?si)<ac:link.*?>.*?<ac:plain-text-link-body>\s*<!\[CDATA\[(.*?)]]>\s*</ac:plain-text-link-body>.*?</ac:link>', '$1' `
                                                                   -replace '(?si)<a\s+[^>]*?href\s*=\s*".*?".*?>\s*(.*?)\s*</a>', '$1' `
                                                                   -replace '(?si)<ac:structured-macro\s+(?:[^>]*?\s+)?ac:name\s*=\s*"jira"(?:\s+[^>]*?)?>.*?<ac:parameter\s+(?:[^>]*?\s+)?ac:name\s*=\s*"key"(?:\s+[^>]*?)?>(.*?)</ac:parameter>.*?</ac:structured-macro>', '$1'
                    $rowHeaderValue = ($processedHeaderContent -replace '<.*?>','').Trim()

                    $rowData[$RowHeaderColumnName] = $rowHeaderValue
                    $dataCellsHtml = $rowHtml.Substring($firstCellMatch.Index + $firstCellMatch.Length) # Get rest of row
                } else { Write-Verbose "Row $rowCount`: No row header (<th>) found as first element." }
            }

            # --- Process Data Cells (<td>) ---
            $cellMatches = [regex]::Matches($dataCellsHtml, '(?si)<td.*?>(.*?)</td>')
            $cellCount = $cellMatches.Count
            Write-Verbose "Row $rowCount`: Found $cellCount data cell(s) (<td>)."

            if ($columnHeaders.Count -gt 0 -and $cellCount -ne $columnHeaders.Count) { Write-Warning "Row $rowCount`: Data cell count ($cellCount) != column header count ($($columnHeaders.Count)). Misalignment likely." }

            $cellIndex = 0
            foreach ($cellMatch in $cellMatches) {
                $headerName = if ($cellIndex -lt $columnHeaders.Count) { $columnHeaders[$cellIndex] } else { "DataColumn$($cellIndex + 1)" }
                $rawCellContent = $cellMatch.Groups[1].Value
                $processedCellContent = $rawCellContent # Start with raw

                # Decode Entities FIRST - critical for accurate macro parsing
                if ($DecodeHtmlEntities) {
                    try { $processedCellContent = [System.Web.HttpUtility]::HtmlDecode($processedCellContent) }
                    catch { Write-Warning "Row $rowCount, Cell $($cellIndex + 1): Error decoding HTML entity: $($_.Exception.Message)." }
                }

                # --- Confluence Macro Parsing ---
                # Apply replacements sequentially. Order might matter in complex cases.
                # 1. Jira Macro (extract key)
                $processedCellContent = $processedCellContent -replace '(?si)<ac:structured-macro\s+(?:[^>]*?\s+)?ac:name\s*=\s*"jira"(?:\s+[^>]*?)?>.*?<ac:parameter\s+(?:[^>]*?\s+)?ac:name\s*=\s*"key"(?:\s+[^>]*?)?>(.*?)</ac:parameter>.*?</ac:structured-macro>', '$1'
                # 2. User Link (extract display name)
                $processedCellContent = $processedCellContent -replace '(?si)<ac:userlink.*?>\s*(.*?)\s*</ac:userlink>', '$1'
                # 3. Confluence Link (extract CDATA body)
                $processedCellContent = $processedCellContent -replace '(?si)<ac:link.*?>.*?<ac:plain-text-link-body>\s*<!\[CDATA\[(.*?)]]>\s*</ac:plain-text-link-body>.*?</ac:link>', '$1'
                # 4. Standard Link (extract display text)
                $processedCellContent = $processedCellContent -replace '(?si)<a\s+[^>]*?href\s*=\s*".*?".*?>\s*(.*?)\s*</a>', '$1'

                # --- Final Cleanup ---
                # Strip remaining simple HTML tags (like <b>, <i>, <span> etc.) and trim
                $finalCellContent = ($processedCellContent -replace '<.*?>','').Trim()

                $rowData[$headerName] = $finalCellContent
                $cellIndex++
            }

             # Pad missing DATA cells based on COLUMN headers
             if ($columnHeaders.Count -gt 0 -and $cellCount -lt $columnHeaders.Count) {
                 for ($i = $cellCount; $i -lt $columnHeaders.Count; $i++) {
                    $headerName = $columnHeaders[$i]; $rowData[$headerName] = $null; Write-Verbose "Row $rowCount`: Padding missing value for column header '$headerName'."
                 }
             }

            # Convert to PSCustomObject
            if ($rowData.Count -gt 0) { $outputObjects.Add([PSCustomObject]$rowData) }
            else { Write-Verbose "Row $rowCount resulted in empty data object, skipping." }
        }

        Write-Verbose "[$((Get-Date).TimeOfDay)] End block finished. Extracted $($outputObjects.Count) rows."
        return $outputObjects
    } # End End block
} # End Function
