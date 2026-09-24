---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentExpand

## SYNOPSIS
Creates a Confluence expand macro (a collapsible section).

## SYNTAX

```
New-ConfluenceContentExpand [[-Title] <String>] [-Content] <String> [-Raw] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentExpand returns the storage-format markup for the Confluence "expand"
macro: a section that is collapsed until the reader clicks its title.
The title is escaped.
The content is escaped and wrapped in a paragraph; use -Raw to insert storage-format markup,
for example a table or code block built with the other New-ConfluenceContent* functions.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentExpand -Title 'Details' -Content (New-ConfluenceContentTable -TableData $rows) -Raw
```

Returns a collapsed "Details" section that contains a table.

## PARAMETERS

### -Title
The text of the clickable title.
Confluence shows "Click here to expand..." when it is empty.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
The body of the section: plain text, or storage-format markup with -Raw.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Raw
Insert -Content as storage-format markup without escaping or wrapping it in a paragraph.

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
