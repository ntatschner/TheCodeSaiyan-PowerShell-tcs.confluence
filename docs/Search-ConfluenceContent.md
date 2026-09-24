---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Search-ConfluenceContent

## SYNOPSIS
Searches Confluence with a CQL query.

## SYNTAX

```
Search-ConfluenceContent [-Cql] <String> [-Limit <Int32>] [-MaxQueryPages <Int16>] [-All] [-Expand <String[]>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
Search-ConfluenceContent runs a Confluence Query Language (CQL) search through the v1 search
API (GET /wiki/rest/api/search?cql=...) and writes each search result to the pipeline.
A
result has content (the page, blog post or attachment), title, excerpt, url and
lastModified properties.

Up to -MaxQueryPages result pages are read; a warning is written when more results are
available.
Use -All to read every result page.

The query is sent as given: quote values with double quotes and escape a backslash or double
quote inside a value with a backslash.

## EXAMPLES

### EXAMPLE 1
```
now("-7d")' -All
```

Returns every page in the DOCS space changed in the last seven days.

### EXAMPLE 2
```
Search-ConfluenceContent -Cql 'label = runbook' | Select-Object title, url
```

Lists the content labelled runbook.

## PARAMETERS

### -Cql
The CQL query, for example: type = page AND space = DOCS AND text ~ "backup".

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Limit
The number of results requested per API call.
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
The maximum number of API result pages to retrieve.
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
Retrieve every API result page (ignores -MaxQueryPages).

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

### -Expand
Properties to expand in the results, for example content.space or content.version.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

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

### System.Management.Automation.PSCustomObject. One object per search result.
## NOTES

## RELATED LINKS
