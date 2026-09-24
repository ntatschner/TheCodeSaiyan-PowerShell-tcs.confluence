---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-ConfluencePageContent

## SYNOPSIS
Gets Confluence pages including their body in the requested format.

## SYNTAX

### ById (Default)
```
Get-ConfluencePageContent -PageId <String> [-ResultsLimit <Int32>] [-MaxQueryPages <Int16>]
 [-ContentType <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### Search
```
Get-ConfluencePageContent [-SpaceKey <String>] [-Title <String>] [-ResultsLimit <Int32>]
 [-MaxQueryPages <Int16>] [-All] [-ContentType <String>] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
Get-ConfluencePageContent retrieves a page by ID, or searches pages by space and title, and asks
Confluence to include the page body in the format given by -ContentType (storage by default)
and writes the page objects to the pipeline.
The body is in body.\<format\>.value.

By ID and exact titles use the v2 pages API (body-format=\<format\>).
A title containing
wildcards (* ?) is searched with CQL through /wiki/rest/api/content/search, which returns v1
content objects; the body is requested with expand=body.\<format\>.

Up to -MaxQueryPages API result pages are read; a warning is written when more results are
available.
Use -All to read every result page.

Before 0.2.0 the pages were wrapped in an object with Results and MultiPage properties.

## EXAMPLES

### EXAMPLE 1
```
(Get-ConfluencePageContent -PageId 123456).body.storage.value
```

Returns the storage-format body of page 123456.

### EXAMPLE 2
```
Get-ConfluencePageContent -SpaceKey DOCS -Title 'R&D [draft]' -ContentType view
```

Returns the "R&D \[draft\]" page in the DOCS space with its rendered HTML body.

## PARAMETERS

### -PageId
The ID of the page to retrieve.

```yaml
Type: String
Parameter Sets: ById
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpaceKey
The space key or numeric space ID to search in.

```yaml
Type: String
Parameter Sets: Search
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
The exact page title, or a title with wildcards (* ?).

```yaml
Type: String
Parameter Sets: Search
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ResultsLimit
The number of results requested per API call when searching.
Default 25.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 25
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxQueryPages
The maximum number of API result pages to retrieve when searching.
Default 3.

```yaml
Type: Int16
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 3
Accept pipeline input: False
Accept wildcard characters: False
```

### -All
Retrieve every API result page when searching (ignores -MaxQueryPages).

```yaml
Type: SwitchParameter
Parameter Sets: Search
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContentType
The body format to return: storage (default), atlas_doc_format, view, export_view,
styled_view, anonymous_export_view or editor.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Storage
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
