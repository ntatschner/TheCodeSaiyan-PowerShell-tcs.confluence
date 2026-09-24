---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluencePageLayout

## SYNOPSIS
Creates a Confluence page layout with one or more sections of one to three columns.

## SYNTAX

### Single (Default)
```
New-ConfluencePageLayout [-LayoutType] <String> [-SectionOne] <String> [[-SectionTwo] <String>]
 [[-SectionThree] <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### Multiple
```
New-ConfluencePageLayout -Section <Hashtable[]> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluencePageLayout returns the storage-format markup for an ac:layout.

For a layout with one section, give -LayoutType and the content of each column with
-SectionOne, -SectionTwo and -SectionThree: single takes one column, two_equal,
two_left_sidebar and two_right_sidebar take two, three_equal and three_with_sidebars take
three.
Passing the wrong number of columns is an error.

For a layout with several sections, pass -Section with one hashtable per section, each with a
LayoutType and a Cells array, for example @{ LayoutType = 'two_equal'; Cells = '\<p\>a\</p\>', '\<p\>b\</p\>' }.

The column content is inserted as given; it is expected to be storage-format markup, such as
the output of the New-ConfluenceContent* functions.

Before 0.2.0 this function returned an object with LayoutType, LayoutXml and ContentSections;
it now returns the markup string like the other builders.

## EXAMPLES

### EXAMPLE 1
```
Left</p>' -SectionTwo '<p>Right</p>'
```

Returns a layout with one two-column section.

### EXAMPLE 2
```
Report</h1>' }, @{ LayoutType = 'two_equal'; Cells = '<p>Left</p>', '<p>Right</p>' }
```

Returns a layout with a full-width section followed by a two-column section.

## PARAMETERS

### -LayoutType
The layout of a single section: single, two_equal, two_left_sidebar, two_right_sidebar,
three_equal or three_with_sidebars.

```yaml
Type: String
Parameter Sets: Single
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SectionOne
The content of the first column.

```yaml
Type: String
Parameter Sets: Single
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SectionTwo
The content of the second column (two- and three-column layouts).

```yaml
Type: String
Parameter Sets: Single
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SectionThree
The content of the third column (three-column layouts).

```yaml
Type: String
Parameter Sets: Single
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Section
Several sections, in order: hashtables with LayoutType and Cells (one string per column).

```yaml
Type: Hashtable[]
Parameter Sets: Multiple
Aliases:

Required: True
Position: Named
Default value: None
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
