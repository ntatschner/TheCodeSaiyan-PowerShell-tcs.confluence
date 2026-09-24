---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-ConfluencePage

## SYNOPSIS
Gets Confluence pages by ID, by space or by title.

## SYNTAX

### AllPages (Default)
```
Get-ConfluencePage [-SpaceKey <String>] [-Search <String>] [-ResultsLimit <Int32>] [-MaxQueryPages <Int16>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### PageId
```
Get-ConfluencePage -PageId <String> [-ResultsLimit <Int32>] [-MaxQueryPages <Int16>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
Get-ConfluencePage retrieves pages through the Confluence v2 pages API.
Use -PageId for a single
page, or -SpaceKey and/or -Search to list pages.
An exact -Search title is sent as a title
filter; a title with wildcards (* ?) is sent as a CQL search through the v1 content API.

Returns the object from Invoke-ConfluenceRequest: the pages are in its Results property.

## EXAMPLES

### EXAMPLE 1
```
(Get-ConfluencePage -PageId 123456).Results
```

Returns page 123456.

### EXAMPLE 2
```
(Get-ConfluencePage -SpaceKey DOCS -Search 'Release notes*').Results | Select-Object id, title
```

Lists the pages in the DOCS space whose title starts with "Release notes".

## PARAMETERS

### -PageId
The ID of the page to retrieve.

```yaml
Type:String
Parameter Sets: PageId
Aliases:
Required: True
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

### -SpaceKey
The key (or numeric ID) of the space whose pages are listed.
Keys are resolved to space IDs.

```yaml
Type:String
Parameter Sets: AllPages
Aliases:
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

### -Search
A page title to match.
Wildcards (* ?) are supported.

```yaml
Type:String
Parameter Sets: AllPages
Aliases:
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

### -ResultsLimit
The number of results requested per API call.
Default 25.

```yaml
Type:
Int32
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: 25
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -MaxQueryPages
The maximum number of API result pages to retrieve.
Default 3.

```yaml
Type:
Int16
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: 3
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

### System.Management.Automation.PSCustomObject with Results and MultiPage properties.
## NOTES

## RELATED LINKS
