---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluencePage

## SYNOPSIS
Creates a Confluence page, or with -Force updates the existing page with the same title.

## SYNTAX

```
New-ConfluencePage [-SpaceKey] <String> [-ParentId] <String> [-Title] <String> [-Status] <String>
 [-Content] <String> [-Force] [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluencePage creates a page under a parent page through the Confluence v2 pages API.
The
body is sent in storage format.

When Confluence reports that a page with the same title already exists and -Force is given,
the existing page with exactly that title (preferring the one under -ParentId) is updated with
the new content as a new version.
Without -Force the conflict is reported as an error.

If Confluence reports any other error but a page with exactly this title under this parent
exists afterwards (Confluence sometimes creates the page and still returns an error), that page
is returned with a warning.

Supports -WhatIf and -Confirm.
Returns the created or updated page.

## EXAMPLES

### EXAMPLE 1
```
Notes</p>'
```

Creates the page "Release 1.2" under page 1000.

### EXAMPLE 2
```
New-ConfluencePage -SpaceKey 98765 -ParentId 1000 -Title 'Daily report' -Status current -Content $html -Force
```

Creates the page, or replaces the content of the existing "Daily report" page.

## PARAMETERS

### -SpaceKey
The numeric ID of the space to create the page in (the v2 API field spaceId).

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
The ID of the parent page.

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

### -Title
The page title.

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

### -Status
The page status: current (published) or draft.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: True
Position: 4Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Content
The page body in Confluence storage format (XHTML), for example built with the
New-ConfluenceContent* functions.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: True
Position: 5Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Force
When a page with the same title exists, update it instead of failing.
This replaces the
content of that page.

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

### System.Management.Automation.PSCustomObject. The created or updated page.
## NOTES

## RELATED LINKS
