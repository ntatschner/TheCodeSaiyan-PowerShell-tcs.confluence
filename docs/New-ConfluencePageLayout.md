---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluencePageLayout

## SYNOPSIS
Creates a Confluence page layout section with one to three columns.

## SYNTAX

```
New-ConfluencePageLayout [-LayoutType] <String> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluencePageLayout returns the storage-format markup for an ac:layout with one section
of the chosen type.
The content for each column is passed through dynamic parameters that
appear once -LayoutType is given: -SectionOne for single, -SectionOne and -SectionTwo for the
two-column layouts, and -SectionOne, -SectionTwo and -SectionThree for the three-column layouts.

Returns an object with LayoutType, LayoutXml (the markup) and ContentSections (the number of
columns).

## EXAMPLES

### EXAMPLE 1
```
Left</p>' -SectionTwo '<p>Right</p>').LayoutXml
```

Returns a two-column layout.

## PARAMETERS

### -LayoutType
The layout: single, two_equal, two_left_sidebar, two_right_sidebar, three_equal or
three_with_sidebars.

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

### System.Management.Automation.PSCustomObject
## NOTES

## RELATED LINKS
