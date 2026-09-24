---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Join-ConfluenceContent

## SYNOPSIS
Joins blocks of Confluence content with a separator.

## SYNTAX

```
Join-ConfluenceContent [-ContentBlocks] <String[]> [[-Separator] <String>] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
Join-ConfluenceContent concatenates storage-format content blocks, such as the output of the
New-ConfluenceContent* functions, with a horizontal rule, line break, space or tab between them.

## EXAMPLES

### EXAMPLE 1
```
Text</p>' -Separator HorizontalRule
```

Returns \<h1\>Intro\</h1\>\<hr /\>\<p\>Text\</p\>.

## PARAMETERS

### -ContentBlocks
The content blocks to join, in order.

```yaml
Type: String[]
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

### -Separator
The separator between blocks: NewLine (\<br /\>, default), HorizontalRule (\<hr /\>),
Space (&nbsp;) or Tab (&emsp;).

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 2Default
Default value: None
Default value: NewLine
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
