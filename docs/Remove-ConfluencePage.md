---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Remove-ConfluencePage

## SYNOPSIS
Deletes a Confluence page.

## SYNTAX

```
Remove-ConfluencePage [-PageId] <String> [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Remove-ConfluencePage deletes the page with the given ID through the Confluence v2 pages API.
The page moves to the space trash.
Because this changes the site, you are asked to confirm
unless you pass -Confirm:$false; -WhatIf shows what would be deleted.

## EXAMPLES

### EXAMPLE 1
```
Remove-ConfluencePage -PageId 123456
```

Asks for confirmation, then deletes page 123456.

### EXAMPLE 2
```
(Get-ConfluencePage -SpaceKey DOCS -Search 'Draft*').Results | Remove-ConfluencePage -WhatIf
```

Shows which draft pages would be deleted without deleting them.

## PARAMETERS

### -PageId
The ID of the page to delete.
Accepts pipeline input by property name (id).

```yaml
Type:String
Parameter Sets:   (All)
Aliases:id
Required: True
Position: 1Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
Accept wildcard characters: False
```

### -WhatIf
Shows what would happen if the cmdlet runs.
The cmdlet is not run.

```yaml
Type:Switch
Parameter Sets:   (All)
Aliases:wi
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

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type:Switch
Parameter Sets:   (All)
Aliases:cf
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

### None.
## NOTES

## RELATED LINKS
