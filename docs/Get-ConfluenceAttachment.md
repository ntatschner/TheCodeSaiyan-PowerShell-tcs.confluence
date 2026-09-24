---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-ConfluenceAttachment

## SYNOPSIS
Gets the attachments of a Confluence page.

## SYNTAX

```
Get-ConfluenceAttachment [-PageId] <String> [[-FileName] <String>] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
Get-ConfluenceAttachment returns the attachments of a page through the Confluence v2 API
(GET /wiki/api/v2/pages/\<id\>/attachments), reading every result page.
Each attachment has
id, title, mediaType, fileSize, version and downloadLink properties.

## EXAMPLES

### EXAMPLE 1
```
Get-ConfluenceAttachment -PageId 123456 | Select-Object title, fileSize
```

Lists the files attached to page 123456.

## PARAMETERS

### -PageId
The ID of the page.
Accepts pipeline input by property name (id).

```yaml
Type: String
Parameter Sets: (All)
Aliases: id

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -FileName
Only return the attachment with this file name.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
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

### System.Management.Automation.PSCustomObject. One object per attachment.
## NOTES

## RELATED LINKS
