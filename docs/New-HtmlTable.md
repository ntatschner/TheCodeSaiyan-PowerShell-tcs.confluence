---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-HtmlTable

## SYNOPSIS
Creates a new HTML table or merges data into an existing HTML table string.

## SYNTAX

### CreateNew (Default)
```
New-HtmlTable [-InputObject] <Object[]> [[-Properties] <String[]>] [-Title <String>] [-CssClass <String>]
 [-Style <Hashtable>] [-HeaderStyle <Hashtable>] [-CellStyle <Hashtable>] [-MergeRows]
 [-MergeColumns <String[]>] [-NullDisplay <String>] [-UseNbspForEmpty] [-PreserveHtml]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### MergeExisting
```
New-HtmlTable [-InputObject] <Object[]> [[-Properties] <String[]>] [-CellStyle <Hashtable>] [-MergeRows]
 [-MergeColumns <String[]>] [-ExistingHtmlTable] <String> [-MergeWithExisting] [-NullDisplay <String>]
 [-UseNbspForEmpty] [-PreserveHtml] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
This advanced function generates or modifies an HTML table string.
In 'CreateNew' mode, it builds a table from input objects with styling and optional row merging.
In 'MergeExisting' mode, it parses an existing HTML table string, detects its columns (best effort),
strips simple HTML tags from detected header names, applies the provided -CellStyle to ALL \<td\> elements
within the existing \<tbody\> (overwriting existing styles), and appends new data rows (InputObject) also
applying the -CellStyle.
Potentially applies row merging relative to the last existing row.

Cell values are HTML-encoded unless -PreserveHtml is given.

## EXAMPLES

### EXAMPLE 1
```
Get-Process | Select-Object -First 5 Name, Id | New-HtmlTable -Title 'Processes' -Style @{ width = '100%' }
```

Creates a table with a caption from five process objects.

### EXAMPLE 2
```
$rows | New-HtmlTable -MergeRows -MergeColumns Region
```

Creates a table in which consecutive rows with the same Region share one Region cell.

### EXAMPLE 3
```
$newRows | New-HtmlTable -ExistingHtmlTable $existingHtml -MergeWithExisting
```

Appends the new rows to the body of an existing table.

## PARAMETERS

### -InputObject
The objects to add as rows.
Accepts pipeline input.

```yaml
Type: Object[]
Parameter Sets:   (All)
Aliases:
Required: True
Position: 1Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Properties
The properties to use as columns, in order.
Defaults to the properties of the first object
(CreateNew) or the columns detected in the existing table (MergeExisting).

```yaml
Type: String[]
Parameter Sets:   (All)
Aliases:
Required: False
Position: 2Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Title
A caption for a new table.

```yaml
Type:String
Parameter Sets: CreateNew
Aliases:
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

### -CssClass
The CSS class of a new table.

```yaml
Type:String
Parameter Sets: CreateNew
Aliases:
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

### -Style
Inline styles for a new table, as a hashtable such as @{ width = '100%' }.

```yaml
Type:Hashtable
Parameter Sets: CreateNew
Aliases:
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

### -HeaderStyle
Inline styles for the header cells of a new table.

```yaml
Type:Hashtable
Parameter Sets: CreateNew
Aliases:
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

### -CellStyle
Inline styles for the data cells.
With -MergeWithExisting the style is also applied to all
existing cells.

```yaml
Type:Hashtable
Parameter Sets:   (All)
Aliases:
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

### -MergeRows
Merge consecutive cells with equal values in -MergeColumns using rowspan.

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

### -MergeColumns
The columns whose equal consecutive values are merged when -MergeRows is given.

```yaml
Type: String[]
Parameter Sets:   (All)
Aliases:
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

### -ExistingHtmlTable
The HTML of an existing table to append rows to.

```yaml
Type:String
Parameter Sets: MergeExisting
Aliases:
Required: True
Position: 2Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -MergeWithExisting
Append the objects to -ExistingHtmlTable instead of creating a new table.

```yaml
Type:Switch
Parameter Sets: MergeExisting
Aliases:
Required: True
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

### -NullDisplay
The text shown for null or empty values.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
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

### -UseNbspForEmpty
Show &nbsp; for null or empty values.

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

### -PreserveHtml
Insert values without HTML-encoding them, so they may contain markup.

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
