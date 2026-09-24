---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Remove-DuplicateConfluencePage

## SYNOPSIS
Removes duplicate Confluence pages that share a title and parent, keeping one.

## SYNTAX

```
Remove-DuplicateConfluencePage [-SpaceKey] <String> [-ParentId] <String> [-PageTitle] <String> [-KeepNewest]
 [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Remove-DuplicateConfluencePage lists the pages in a space, finds the pages with the given title
under the given parent and, when there is more than one, deletes all but one.
By default the
newest page (highest version number, then latest modification) is kept; use
-KeepNewest:$false to keep the oldest.

Every deletion asks for confirmation unless you pass -Confirm:$false; -WhatIf shows what would
be deleted.
The kept page is returned.
Only the pages returned by Get-ConfluencePage for the
space (up to three API result pages) are considered.

Called as Remove-DuplicateConfluencePages (the name used before 0.1.0) it still works through
an alias.

## EXAMPLES

### EXAMPLE 1
```
Remove-DuplicateConfluencePage -SpaceKey DOCS -ParentId 1000 -PageTitle 'Weekly report' -WhatIf
```

Shows which duplicate "Weekly report" pages under page 1000 would be deleted.

### EXAMPLE 2
```
Remove-DuplicateConfluencePage -SpaceKey DOCS -ParentId 1000 -PageTitle 'Weekly report' -KeepNewest:$false -Confirm:$false
```

Deletes all but the oldest duplicate without prompting.

## PARAMETERS

### -SpaceKey
The key or numeric ID of the space to search.

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

### -ParentId
The ID of the parent page of the duplicates.

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

### -PageTitle
The exact title of the duplicated page.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: True
Position: 3Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -KeepNewest
Keep the newest page (default).
Use -KeepNewest:$false to keep the oldest page instead.

```yaml
Type:Switch
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: False
Accept pipeline input: False
input:False
Accept pipeline input: False
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

### System.Management.Automation.PSCustomObject. The page that was kept.
## NOTES

## RELATED LINKS
