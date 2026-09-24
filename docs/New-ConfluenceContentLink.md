---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentLink

## SYNOPSIS
Creates a hyperlink, or turns every URL in a block of text into a hyperlink.

## SYNTAX

### Url (Default)
```
New-ConfluenceContentLink [-LinkText <String>] -Url <String> [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

### TextBlock
```
New-ConfluenceContentLink -TextBlock <String> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
With -Url, New-ConfluenceContentLink returns an \<a\> element that shows -LinkText (or the URL
when no text is given).
With -TextBlock, every address with a scheme (http://, https:// or
mailto:) found in the text is replaced by a link to itself; file names such as report.pdf and
bare e-mail addresses are left as text.

The URL, the link text and the text block are escaped (& \< \> and quotes), so the result is
always well-formed storage format.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentLink -Url 'https://example.com' -LinkText 'Example'
```

Returns \<a href='https://example.com'\>Example\</a\>.

### EXAMPLE 2
```
New-ConfluenceContentLink -TextBlock 'See https://example.com/docs for details.'
```

Returns the sentence with the address converted to a link.

## PARAMETERS

### -LinkText
The text shown for the link.
Defaults to the URL.

```yaml
Type: String
Parameter Sets: Url
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Url
The address to link to.

```yaml
Type: String
Parameter Sets: Url
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextBlock
Text in which every http, https or mailto address is converted to a link.

```yaml
Type: String
Parameter Sets: TextBlock
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
