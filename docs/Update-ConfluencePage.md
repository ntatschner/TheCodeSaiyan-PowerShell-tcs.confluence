---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Update-ConfluencePage

## SYNOPSIS
Replaces the title and body of a Confluence page.

## SYNTAX

```
Update-ConfluencePage [-PageId] <String> [[-SpaceKey] <String>] [-Title] <String> [[-Status] <String>]
 [-Content] <String> [-Version] <Int32> [[-VersionMessage] <String>] [-ProgressAction <ActionPreference>]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Update-ConfluencePage sends a new version of a page through the Confluence v2 pages API.
The
body is sent in storage format.
Confluence requires the new version number, which is the
current version number plus one.

Supports -WhatIf and -Confirm.
Returns the updated page.

## EXAMPLES

### EXAMPLE 1
```
$page = (Get-ConfluencePage -PageId 123456).Results[0]
Update-ConfluencePage -PageId $page.id -Title $page.title -Content '<p>New body</p>' -Version ($page.version.number + 1)
```

Replaces the body of page 123456.

## PARAMETERS

### -PageId
The ID of the page to update.

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

### -SpaceKey
The numeric space ID to move the page to.
Leave empty to keep the page in its space.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
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
The page status: current (default) or draft.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 4Default
Default value: None
Default value: Current
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Content
The page body in Confluence storage format (XHTML).

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

### -Version
The new version number: the page's current version number plus one.

```yaml
Type:
Int32
Parameter Sets:   (All)
Aliases:
Required: True
Position: 6Default
Default value: None
Default value: 0
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -VersionMessage
The version comment shown in the page history.
Default "Programmatically Updated".

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 7Default
Default value: None
Default value: ProgrammaticallyUpdated
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

### System.Management.Automation.PSCustomObject. The updated page.
## NOTES

## RELATED LINKS
