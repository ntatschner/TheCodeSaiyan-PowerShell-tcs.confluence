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
ConvertTo-ConfluenceHTML performs a lightweight, regex-based Markdown conversion.
Supported:
headings (# to ######), **bold**, *italic*, "- " bullet lists, fenced code blocks (\`\`\`), which
may span several lines and whose content is HTML-encoded and left untouched by the other
rules, and table rows (| a | b |), each converted to a single-cell table row.

It is not a full Markdown parser; nested lists, links, images and inline code are left as they
are.

## EXAMPLES

### EXAMPLE 1
```
ConvertTo-ConfluenceHTML -InputFormat Markdown -InputContent "# Title`n- one`n- two"
```

Returns \<h1\>Title\</h1\> followed by \<ul\>\<li\>one\</li\>\<li\>two\</li\>\</ul\>.

## PARAMETERS

### -InputContent
The Markdown text to convert.

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

### -InputFormat
The format of the input.
Only Markdown is supported.

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
