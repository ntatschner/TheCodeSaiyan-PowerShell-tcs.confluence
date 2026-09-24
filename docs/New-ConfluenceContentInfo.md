---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentInfo

## SYNOPSIS
Creates a Confluence info, tip, note or warning panel macro.

## SYNTAX

```
New-ConfluenceContentInfo [-Title] <String> [-Content] <String> [[-Type] <String>] [-Raw]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentInfo returns the storage-format markup for the Confluence info, tip, note
or warning macro (\<ac:structured-macro ac:name="info"\> with a title parameter and a
rich-text body).
Confluence has no "error" panel macro, so -Type error produces the warning
macro.

The title is always escaped.
The content is escaped and wrapped in a paragraph; use -Raw to
insert storage-format markup (for example the output of other New-ConfluenceContent*
functions) as the body instead.

Before 0.2.0 this function returned an AUI message div, which Confluence does not render as a
panel.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentInfo -Title 'Heads up' -Content 'Maintenance on Friday.' -Type warning
```

Returns a warning panel.

### EXAMPLE 2
```
New-ConfluenceContentInfo -Title 'Steps' -Content (New-ConfluenceContentCodeBlock -Content 'Restart-Service web') -Raw
```

Returns an info panel whose body is a code block.

## PARAMETERS

### -Title
The title of the panel.
An empty title shows no title.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
The body of the panel: plain text, or storage-format markup with -Raw.

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

### -Type
The panel type: info (default), tip, note, warning or error (rendered as warning).

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: Info
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
