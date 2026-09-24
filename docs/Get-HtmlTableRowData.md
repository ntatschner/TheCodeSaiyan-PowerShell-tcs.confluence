---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-HtmlTableRowData

## SYNOPSIS
Extracts row data from an HTML table string into PowerShell objects, attempting to parse common Confluence macros.

## SYNTAX

```
Get-HtmlTableRowData [-HtmlContent] <String> [-TableIndex <Int32>] [-NoHeader] [-NoRowHeader]
 [-RowHeaderColumnName <String>] [-DecodeHtmlEntities <Boolean>] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
This function parses an HTML string containing one or more tables and attempts
to extract the data from the cells (\<td\>) within the body (\<tbody\>) of a specified table.
It uses regular expressions for parsing, which has limitations with complex or
malformed HTML.

It attempts to detect column headers (\<th\> in \<thead\>) and optionally row headers
(first \<th\> in a \<tbody\> row).

Crucially, it includes logic to specifically parse common Confluence macros found within
cell content before general tag stripping:
- User Links (\<ac:userlink\>): Replaces with the user's display name.
- Confluence Links (\<ac:link\>): Replaces with the link's display text.
- Standard Links (\<a\>): Replaces with the link's display text.
- Jira Issues (\<ac:structured-macro name="jira"\>): Replaces with the Jira issue key.

## EXAMPLES

### EXAMPLE 1
```
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
```

# Expected Output (may vary slightly based on exact macro rendering):
# Task     Assignee    Status      Related Link
# ----     --------    ------      ------------
# PROJ-123 John Doe    Done        Project Docs
# PROJ-456 Alice Smith In Progress External Site

### EXAMPLE 2
```
# Example 2: No column headers, but row headers and macros
$htmlRowMacro = @"
<table>
    <tr><th>PROJ-123</th><td><ac:userlink>John Doe</ac:userlink></td><td>Done</td></tr>
    <tr><th>PROJ-456</th><td><ac:userlink>Alice Smith</ac:userlink></td><td>WIP</td></tr>
</table>
"@
Get-HtmlTableRowData -HtmlContent $htmlRowMacro -NoHeader -RowHeaderColumnName "JiraKey"
```

# Output:
# JiraKey  Column1     Column2
# -------  -------     -------
# PROJ-123 John Doe    Done
# PROJ-456 Alice Smith WIP

## PARAMETERS

### -HtmlContent
A string containing the HTML source code that includes the table(s).
Often from a Confluence export or API.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: True
Position: 1Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -TableIndex
The zero-based index of the table to extract data from if multiple tables exist
in the HtmlContent.
Defaults to 0 (the first table found).

```yaml
Type:
Int32
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: 0
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -NoHeader
Switch parameter.
If specified, the function will not attempt to read column headers
from \<thead\>\<th\> tags.
It will use default property names like "Column1", "Column2", etc.,
for the data cells (\<td\>).
Row header detection still occurs unless -NoRowHeader is also specified.

```yaml
Type:Switch
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: False
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -NoRowHeader
Switch parameter.
If specified, the function will not treat the first \<th\> element in a
\<tbody\> row as a special row header.

```yaml
Type:Switch
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: False
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -RowHeaderColumnName
The property name to use in the output objects for the data extracted from row headers
(first \<th\> cell in a \<tbody\> row).
Defaults to "RowHeader".
Must be a valid simple variable name.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: RowHeader
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -DecodeHtmlEntities
Boolean parameter.
When $true (the default), attempts to decode HTML entities (like &amp;, &lt;, &nbsp;)
found within header and table cell data using \[System.Net.WebUtility\]::HtmlDecode.
Specify -DecodeHtmlEntities $false to keep entities as they are.

```yaml
Type:Boolean
Parameter Sets:   (All)
Aliases:Decode
Required: False
Position:Named
Default value: None
Default value: None
Default value: True
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -ProgressAction
{{ Fill ProgressAction Description }}

```yaml
Type:ActionPreference
Parameter Sets:   (All)
Aliases:proga
Required: False
Position:Named
Default value: None
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### PSCustomObject[] - An array of PowerShell custom objects, where each object represents a row.
###                  Returns $null if parsing fails at a critical step. Returns @() if no data rows found.
## NOTES
Author: AI Assistant
Date: 2023-10-27
- WARNING: Regex-based parsing of Confluence Storage Format / HTML is VERY fragile.
It relies on specific
  tag structures observed in common exports.
Changes in Confluence versions or complex macro usage WILL break this.
  Consider dedicated Confluence API clients or libraries that handle storage format conversion if robustness is critical.
- Macro parsing attempts to extract the most common user-visible text (display names, link text, Jira keys).
- Assumes standard table structure.
Row headers are assumed to be the first \<th\> in a \<tbody\> row.
- Basic HTML tags *remaining after macro processing* are stripped.

## RELATED LINKS
