---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Remove-ConfluencePageLabel

## SYNOPSIS
Removes labels from a Confluence page.

## SYNTAX

```
Remove-ConfluencePageLabel [-PageId] <String> [-Label] <String[]> [-ProgressAction <ActionPreference>]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Remove-ConfluencePageLabel removes one or more labels from a page through the Confluence v1
API (DELETE /wiki/rest/api/content/\<id\>/label?name=\<label\>); the v2 API has no endpoint to
remove labels.
Supports -WhatIf and -Confirm.

## EXAMPLES

### EXAMPLE 1
```
Remove-ConfluencePageLabel -PageId 123456 -Label draft
```

Removes the draft label from page 123456.

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

### -Label
The label names to remove.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
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

### None.
## NOTES

## RELATED LINKS
