---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Add-ConfluencePageLabel

## SYNOPSIS
Adds labels to a Confluence page.

## SYNTAX

```
Add-ConfluencePageLabel [-PageId] <String> [-Label] <String[]> [-ProgressAction <ActionPreference>] [-WhatIf]
 [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Add-ConfluencePageLabel adds one or more global labels to a page through the Confluence v1 API
(POST /wiki/rest/api/content/\<id\>/label); the v2 API has no endpoint to add labels.
Labels
that the page already has are left as they are.
Confluence stores labels in lower case and
they cannot contain spaces.
Returns the labels of the page after the change.

Supports -WhatIf and -Confirm.

## EXAMPLES

### EXAMPLE 1
```
Add-ConfluencePageLabel -PageId 123456 -Label runbook, ops
```

Adds the labels runbook and ops to page 123456.

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
The label names to add.

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

### System.Management.Automation.PSCustomObject. The labels of the page.
## NOTES

## RELATED LINKS
