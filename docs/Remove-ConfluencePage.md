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
Remove-ConfluencePage [-PageId] <String> [-Purge] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Remove-ConfluencePage deletes the page with the given ID through the Confluence v2 pages API.
The page moves to the space trash; with -Purge it is then permanently deleted from the trash
(this needs space admin permission and cannot be undone).
Because this changes the site, you
are asked to confirm unless you pass -Confirm:$false; -WhatIf shows what would be deleted.

## EXAMPLES

### EXAMPLE 1
```
Remove-ConfluencePage -PageId 123456
```

Asks for confirmation, then deletes page 123456.

### EXAMPLE 2
```
Get-ConfluencePage -SpaceKey DOCS -Search 'Draft*' | Remove-ConfluencePage -WhatIf
```

Shows which draft pages would be deleted without deleting them.

### EXAMPLE 3
```
Remove-ConfluencePage -PageId 123456 -Purge -Confirm:$false
```

Deletes page 123456 and purges it from the trash.

## PARAMETERS

### -PageId
The ID of the page to delete.
Accepts pipeline input by property name (id), for example the
pages written by Get-ConfluencePage.

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

### -Purge
Permanently delete the page: it is moved to the trash (if it is not there already) and then
purged with DELETE /wiki/api/v2/pages/\<id\>?purge=true.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### None.
## NOTES

## RELATED LINKS
