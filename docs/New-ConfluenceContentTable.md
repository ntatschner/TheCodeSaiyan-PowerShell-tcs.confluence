---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentTable

## SYNOPSIS
Creates an HTML table for a Confluence page from a collection of objects.

## SYNTAX

```
New-ConfluenceContentTable [-TableData] <Array> [[-TableType] <String>] [[-TableTypeStyle] <String>]
 [-NoHeader] [[-HeaderStringFormatting] <String[]>] [[-HeaderAlignmentFormatting] <String>] [-VerticalHeader]
 [[-CellStringFormatting] <String[]>] [[-CellAlignmentFormatting] <String>] [[-FirstCellHeaderFormat] <String>]
 [[-FirstCellStringFormatting] <String[]>] [[-FirstCellAlignmentFormatting] <String>] [-Raw]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentTable converts objects to a table: the property names of the first object
become the header row and every object becomes a row.
All objects must be of the same type and
have the same property names.
Hashtables and ordered dictionaries are also accepted as rows;
their keys are the columns.

Cell values that are collections, dictionaries or objects with several properties are rendered
as nested tables, up to three levels deep (deeper values are shown as text).
In the data cells (all but the first column) addresses with a scheme
(http://, https:// or mailto:) are converted to links; file names such as report.pdf and bare
e-mail addresses are not.
Header names and cell values are escaped, so the result is always
well-formed storage format; use -Raw to insert cell values that already are storage-format
markup.
An empty collection returns an empty string.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentTable -TableData @([ordered]@{ Service = 'web'; Docs = 'https://contoso.com/web' })
```

Returns a table from an ordered dictionary; the address becomes a link.

### EXAMPLE 2
```
New-ConfluenceContentTable -TableData @([pscustomobject]@{ Name = 'web'; State = (New-ConfluenceContentStatus -Text 'OK' -Colour Green) }) -Raw
```

Returns a table whose State cells contain status lozenges.

### EXAMPLE 3
```
$rows = Get-Process | Select-Object -First 5 Name, Id
New-ConfluenceContentTable -TableData $rows -HeaderStringFormatting Bold -VerticalHeader
```

Returns a table with a bold header row and the process names as row headers.

## PARAMETERS

### -TableData
The objects (or hashtables / ordered dictionaries) to show in the table.

```yaml
Type: Array
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableType
The CSS class(es) of the table, for example Wrapped or "Relative Table".

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Wrapped
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableTypeStyle
A single inline style, for example "Width: 100%" (default), "Width: 600px" or
"Border: 1px solid black".

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: Width: 100%
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Do not add the header row.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderStringFormatting
Formatting for the header cells: Bold, Italic, Underline and/or Strikethrough.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderAlignmentFormatting
Text alignment of the header cells: Left (default), Center or Right.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: Left
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalHeader
Render the first cell of every row as a row header (\<th scope='row'\>).

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -CellStringFormatting
Formatting for the data cells (all but the first column): Bold, Italic, Underline and/or
Strikethrough.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CellAlignmentFormatting
Text alignment of the data cells: Left (default), Center or Right.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: Left
Accept pipeline input: False
Accept wildcard characters: False
```

### -FirstCellHeaderFormat
Wrap the first cell of each row in a heading: 0 (no heading, default) or 1 to 6.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -FirstCellStringFormatting
Formatting for the first cell of each row: Bold, Italic, Underline and/or Strikethrough.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FirstCellAlignmentFormatting
Text alignment of the first cell of each row.
Defaults to -CellAlignmentFormatting.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Raw
Insert cell values as given, without escaping them or converting addresses to links.
Use it
when the values are storage-format fragments, for example links or status lozenges built with
the other New-ConfluenceContent* functions.
Header names are still escaped.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProgressAction
{{ Fill ProgressAction Description }}

```yaml
Type: ActionPreference
Parameter Sets: (All)
Aliases: proga

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.String
## NOTES

## RELATED LINKS
