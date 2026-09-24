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
New-ConfluenceContentHeader [-Header] <String> [-Level] <Int32> [[-StringFormatting] <String[]>] [-Raw]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentHeader returns a heading element with optional bold, italic, underline or
strikethrough formatting.
Formatting tags are nested correctly.
The header text is escaped
(& \< \> become entities) so it always produces valid storage format; use -Raw to insert
markup that you have built yourself.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentHeader -Header 'R&D' -Level 2
```

Returns \<h2\>R&amp;D\</h2\>.

### EXAMPLE 2
```
New-ConfluenceContentHeader -Header 'Summary' -Level 2 -StringFormatting Bold, Italic
```

Returns \<h2\>\<strong\>\<em\>Summary\</em\>\</strong\>\</h2\>.

## PARAMETERS

### -Header
The heading text.

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

### -Level
The heading level, 1 to 6.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -StringFormatting
Formatting to apply: Bold, Italic, Underline and/or Strikethrough.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Raw
Insert -Header as given, without escaping.
Only use it with trusted, well-formed markup.

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
