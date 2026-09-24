---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-ConfluenceSpace

## SYNOPSIS
Gets Confluence spaces, optionally filtered by name.

## SYNTAX

### AllSpaces (Default)
```
Get-ConfluenceSpace [-Search <String>] [-ResultsLimit <Int32>] [-MaxQueryPages <Int16>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### SpaceId
```
Get-ConfluenceSpace -SpaceId <String> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
Get-ConfluenceSpace lists the spaces visible to the account in the Confluence context, or
returns one space by ID.
-Search filters the listed spaces by name on the client side and
supports wildcards (* ?).

Called as Get-ConfluenceSpaces (the name used before 0.1.0) it still works through an alias.

## EXAMPLES

### EXAMPLE 1
```
Get-ConfluenceSpace | Select-Object id, key, name
```

Lists the spaces.

### EXAMPLE 2
```
Get-ConfluenceSpace -Search 'Engineering*'
```

Lists the spaces whose name starts with "Engineering".

## PARAMETERS

### -SpaceId
The numeric ID of the space to retrieve.
Returns the object from Invoke-ConfluenceRequest.

```yaml
Type: String
Parameter Sets: SpaceId
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Search
A space name filter, for example 'Team*'.
Matching is case-insensitive.

```yaml
Type: String
Parameter Sets: AllSpaces
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResultsLimit
The number of spaces requested per API call.
Default 50.

```yaml
Type: Int32
Parameter Sets: AllSpaces
Aliases:

Required: False
Position: Named
Default value: 50
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxQueryPages
The maximum number of API result pages to retrieve.
Default 3.

```yaml
Type: Int16
Parameter Sets: AllSpaces
Aliases:

Required: False
Position: Named
Default value: 3
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

### System.Management.Automation.PSCustomObject
## NOTES

## RELATED LINKS
