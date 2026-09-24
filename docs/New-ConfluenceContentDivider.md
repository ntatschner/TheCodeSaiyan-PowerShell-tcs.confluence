---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentDivider

## SYNOPSIS
Creates a divider (horizontal rule or blank paragraph) for a Confluence page.

## SYNTAX

```
New-ConfluenceContentDivider [[-Type] <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentDivider returns an \<hr\> element, optionally with a style class, or an empty
paragraph for vertical space.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentDivider -Type space
```

Returns \<p\>&nbsp;\</p\>.

## PARAMETERS

### -Type
The divider type: line (default), space, default, dashed, dotted, double or gradient.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 1Default
Default value: None
Default value: Line
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
