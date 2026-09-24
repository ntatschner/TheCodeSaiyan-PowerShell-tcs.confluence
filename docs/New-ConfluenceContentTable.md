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
 [[-FirstCellStringFormatting] <String[]>] [[-FirstCellAlignmentFormatting] <String>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentTable converts objects to a table: the property names of the first object
become the header row and every object becomes a row.
All objects must be of the same type and
have the same property names.

Cell values that are collections or objects with several properties are rendered as nested
tables.
Text that contains a web address is converted to links.
Values are inserted as given,
so they may contain markup.
An empty collection returns an empty string.

## EXAMPLES

### EXAMPLE 1
```
$rows = Get-Process | Select-Object -First 5 Name, Id
New-ConfluenceContentTable -TableData $rows -HeaderStringFormatting Bold -VerticalHeader
```

Returns a table with a bold header row and the process names as row headers.

## PARAMETERS

### -TableData
The objects to show in the table.

```yaml
Type:Array
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

### -TableType
The CSS class(es) of the table, for example Wrapped or "Relative Table".

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 2Default
Default value: None
Default value: Wrapped
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -TableTypeStyle
A single inline style, for example "Width: 100%" (default), "Width: 600px" or
"Border: 1px solid black".

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 3Default
Default value: None
Default value: Width: 100%
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -NoHeader
Do not add the header row.

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

### -HeaderStringFormatting
Formatting for the header cells: Bold, Italic, Underline and/or Strikethrough.

```yaml
Type: String[]
Parameter Sets:   (All)
Aliases:
Required: False
Position: 4Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -HeaderAlignmentFormatting
Text alignment of the header cells: Left (default), Center or Right.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 5Default
Default value: None
Default value: Left
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -VerticalHeader
Render the first cell of every row as a row header (\<th scope='row'\>).

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

### -CellStringFormatting
Formatting for the data cells (all but the first column): Bold, Italic, Underline and/or
Strikethrough.

```yaml
Type: String[]
Parameter Sets:   (All)
Aliases:
Required: False
Position: 6Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -CellAlignmentFormatting
Text alignment of the data cells: Left (default), Center or Right.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 7Default
Default value: None
Default value: Left
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -FirstCellHeaderFormat
Wrap the first cell of each row in a heading: 0 (no heading, default) or 1 to 6.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 8Default
Default value: None
Default value: 0
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -FirstCellStringFormatting
Formatting for the first cell of each row: Bold, Italic, Underline and/or Strikethrough.

```yaml
Type: String[]
Parameter Sets:   (All)
Aliases:
Required: False
Position: 9Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -FirstCellAlignmentFormatting
Text alignment of the first cell of each row.
Defaults to -CellAlignmentFormatting.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 10Default
Default value: None
Default value: None
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

### System.String
## NOTES

## RELATED LINKS
