---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Add-ConfluenceAttachment

## SYNOPSIS
Uploads files as attachments to a Confluence page.

## SYNTAX

```
Add-ConfluenceAttachment [-PageId] <String> [-Path] <String[]> [[-Comment] <String>] [-NotifyWatchers]
 [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Add-ConfluenceAttachment uploads one or more files to a page through the Confluence v1 API
(PUT /wiki/rest/api/content/\<id\>/child/attachment, multipart/form-data with the
X-Atlassian-Token: no-check header); the v2 API has no upload endpoint.
When the page already
has an attachment with the same file name, a new version of that attachment is created.
Returns the created or updated attachments.

Supports -WhatIf and -Confirm.

## EXAMPLES

### EXAMPLE 1
```
Add-ConfluenceAttachment -PageId 123456 -Path ./report.pdf -Comment 'Nightly report'
```

Attaches report.pdf to page 123456, or adds a new version of it.

### EXAMPLE 2
```
Get-ChildItem ./out/*.png | Add-ConfluenceAttachment -PageId 123456
```

Attaches every PNG file in ./out.

## PARAMETERS

### -PageId
The ID of the page to attach the files to.

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

### -Path
The files to upload.
Accepts pipeline input (for example from Get-ChildItem).

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: FullName

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Comment
A comment stored with each attachment version.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotifyWatchers
Notify the page watchers.
By default the upload is a minor edit and watchers are not notified.

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

### -WhatIf
Shows what would happen if the cmdlet runs.
The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
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

### System.Management.Automation.PSCustomObject. One object per attachment.
## NOTES

## RELATED LINKS
