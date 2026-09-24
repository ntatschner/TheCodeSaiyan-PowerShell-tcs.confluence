---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentInfo

## SYNOPSIS
Creates an info, tip, note, warning or error message block.

## SYNTAX

```
New-ConfluenceContentInfo [-Title] <String> [-Content] <String> [[-Type] <String>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentInfo returns an AUI message block (a div with the aui-message classes)
with a title and content.
Title and content are inserted as given, so they may contain markup.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentInfo -Title 'Heads up' -Content 'Maintenance on Friday.' -Type warning
```

Returns a warning message block.

## PARAMETERS

### -Title
The title of the block.

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

### -Content
The body of the block.

```yaml
Type:String
Parameter Sets:   (All)
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

### -Type
The block type: info (default), tip, note, warning or error.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 3Default
Default value: None
Default value: Info
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
