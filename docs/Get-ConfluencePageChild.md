---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-ConfluencePageChild

## SYNOPSIS
Gets the child pages of a Confluence page, or with -Recurse all pages below it.

## SYNTAX

```
Get-ConfluencePageChild [-PageId] <String> [-Recurse] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
Get-ConfluencePageChild returns the child pages of a page through the Confluence v2 API
(GET /wiki/api/v2/pages/\<id\>/children), reading every result page.
With -Recurse the children
of every child are read as well, level by level, so every descendant page is returned (parents
before their children).
Each page gets a parentId property if the API did not return one.

## EXAMPLES

### EXAMPLE 1
```
Get-ConfluencePageChild -PageId 123456 | Select-Object id, title
```

Lists the direct child pages of page 123456.

### EXAMPLE 2
```
Get-ConfluencePageChild -PageId 123456 -Recurse | Measure-Object
```

Counts every page below page 123456.

## PARAMETERS

### -PageId
The ID of the parent page.
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

### -Recurse
Return all descendant pages, not only the direct children.

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

### System.Management.Automation.PSCustomObject. One object per page.
## NOTES

## RELATED LINKS
