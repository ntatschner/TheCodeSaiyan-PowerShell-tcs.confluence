---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentTOC

## SYNOPSIS
Creates a Confluence table of contents macro.

## SYNTAX

```
New-ConfluenceContentTOC [-HorizontalList] [[-BulletPointStyle] <String>] [[-HeadersFromLevel] <Int32>]
 [[-HeadersToLevel] <Int32>] [-IncludeSectionNumbers] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentTOC returns the storage-format markup for the Confluence "toc" macro,
built from the headings on the page.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentTOC -HeadersFromLevel 2 -HeadersToLevel 3 -IncludeSectionNumbers
```

Returns a numbered table of contents of the h2 and h3 headings.

## PARAMETERS

### -HorizontalList
Show the table of contents as a horizontal (flat) list instead of a vertical list.

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

### -BulletPointStyle
The list style: None (default), Mixed (Confluence default bullets), Bullet (disc), Circle,
Square or Number (decimal).

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 1Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -HeadersFromLevel
The lowest heading level to include, 1 to 6.
Default 1.

```yaml
Type:
Int32
Parameter Sets:   (All)
Aliases:
Required: False
Position: 2Default
Default value: None
Default value: 1
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -HeadersToLevel
The highest heading level to include, 1 to 6.
Default 6.

```yaml
Type:
Int32
Parameter Sets:   (All)
Aliases:
Required: False
Position: 3Default
Default value: None
Default value: 6
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -IncludeSectionNumbers
Number the entries as an outline (1, 1.1, 1.2 ...).

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
