---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentHeader

## SYNOPSIS
Creates a heading (h1 to h6) for a Confluence page.

## SYNTAX

```
New-ConfluenceContentHeader [-Header] <String> [-Level] <Int32> [[-StringFormatting] <String[]>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentHeader returns a heading element with optional bold, italic, underline or
strikethrough formatting.
Formatting tags are nested correctly.
The header text is inserted as
given, so it may contain markup.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentHeader -Header 'Summary' -Level 2 -StringFormatting Bold, Italic
```

Returns \<h2\>\<strong\>\<em\>Summary\</em\>\</strong\>\</h2\>.

## PARAMETERS

### -Header
The heading text.

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

### -Level
The heading level, 1 to 6.

```yaml
Type:
Int32
Parameter Sets:   (All)
Aliases:
Required: True
Position: 2Default
Default value: None
Default value: 0
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -StringFormatting
Formatting to apply: Bold, Italic, Underline and/or Strikethrough.

```yaml
Type: String[]
Parameter Sets:   (All)
Aliases:
Required: False
Position: 3Default
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
