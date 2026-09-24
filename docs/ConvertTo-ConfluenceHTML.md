---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# ConvertTo-ConfluenceHTML

## SYNOPSIS
Converts simple Markdown to HTML for a Confluence page body.

## SYNTAX

```
ConvertTo-ConfluenceHTML [-InputContent] <String> [-InputFormat] <String> [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
ConvertTo-ConfluenceHTML performs a lightweight, line-based Markdown conversion.
Supported:
headings (# to ######), **bold**, *italic*, bullet lists ("- " or "* "; consecutive items share
one list), fenced code blocks (\`\`\`), whose content is kept exactly and left untouched by the
other rules, and tables: consecutive | a | b | lines become one table, a |---|---| separator
row is skipped and marks the row above it as the header row.

All text is escaped (& \< \>), so the result is well-formed storage format and HTML in the
Markdown is shown as text.
Other lines are passed through as escaped text, one per line.

It is not a full Markdown parser; nested lists, links, images and inline code are left as
text.

## EXAMPLES

### EXAMPLE 1
```
ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# Title`n- one`n- two"
```

Returns \<h1\>Title\</h1\> followed by \<ul\>\<li\>one\</li\>\<li\>two\</li\>\</ul\>.

### EXAMPLE 2
```
ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "| Name | Count |`n|---|---|`n| a | 1 |"
```

Returns one table with a header row (Name, Count) and one data row.

## PARAMETERS

### -InputContent
The Markdown text to convert.

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

### -InputFormat
The format of the input.
Only Markdown is supported.

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
