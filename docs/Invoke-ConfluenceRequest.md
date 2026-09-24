---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Invoke-ConfluenceRequest

## SYNOPSIS
Sends a request to the Confluence REST API using the current Confluence context.

## SYNTAX

```
Invoke-ConfluenceRequest [-Method] <String> [[-URIPath] <String>] [[-Resource] <String>]
 [[-ApiVersion] <Int32>] [[-Id] <String>] [-RawPath] [[-Body] <String>] [[-Query] <Hashtable>]
 [[-MaxQueryPages] <Int16>] [-All] [[-Search] <String>] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
Invoke-ConfluenceRequest is the transport used by every tcs.confluence REST command.
It builds
the endpoint from -Resource (or an explicit -URIPath) on the site set by Set-ConfluenceContext,
appends query parameters and an optional CQL title search, sends the request with the stored
credential and follows "next" pagination links for GET requests up to -MaxQueryPages pages
(or all pages with -All).

Behaviour that is applied automatically:
- -Resource uses the API version given by -ApiVersion or, when that is not given, the version
  chosen with Set-ConfluenceContext -ApiVersion (v2 by default).
- A -Search against the page or content collection is sent to the v1 CQL search endpoint
  /wiki/rest/api/content/search as cql=title = "..." (or title ~ "..." with wildcards * ?).
  Page searches also add type = page.
CQL values are escaped (backslash, then quote).
- A spaceKey query value is resolved to the numeric space ID and sent as the documented
  space-id parameter for v2 pages (keys are looked up once per session and cached), or added
  to the CQL query (space = "KEY" or space.id = 123) for searches.
A key that cannot be
  resolved is reported as an error and no request is sent.
- The short paths /pages, /content and /spaces are mapped to their full API paths.
- A 429 Too Many Requests response is retried, honouring the Retry-After header.
- Pagination links pointing to a different host are not followed, so the credential is only
  ever sent to the Confluence site in the context.
- When -MaxQueryPages stops the paging while more results are available, a warning is written.

HTTP error responses are reported as errors that include the status and the Confluence error
titles; use -ErrorAction Stop to turn them into terminating errors.

Returns an object with a Results array (all collected results) and a MultiPage flag.

## EXAMPLES

### EXAMPLE 1
```
Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = 'DOCS'; limit = 50 }
```

Returns up to three pages of results for the pages in the DOCS space (sent as space-id).

### EXAMPLE 2
```
Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'Release*' -All
```

Searches every page whose title matches "Release*" through /wiki/rest/api/content/search.

### EXAMPLE 3
```
Invoke-ConfluenceRequest -Method DELETE -Resource pages -Id 123456 -ErrorAction Stop
```

Deletes page 123456 and throws if Confluence reports an error.

## PARAMETERS

### -Method
The HTTP method: GET, POST, PUT or DELETE.

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

### -URIPath
An explicit API path such as /wiki/api/v2/pages.
Used as-is unless -Resource is given without
-URIPath.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Resource
A resource shortcut used to build the path: pages, content or spaces.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ApiVersion
The API version used with -Resource: 2 (/wiki/api/v2) or 1 (/wiki/rest/api).
When not given,
the version from Set-ConfluenceContext -ApiVersion is used.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -Id
A resource ID appended to the path built from -Resource, for example a page ID.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RawPath
Do not build the path from -Resource or map short paths; use -URIPath exactly as given.

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

### -Body
The JSON request body for POST and PUT requests.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Query
A hashtable of query-string parameters, for example @{ limit = 25; spaceKey = 'DOCS' }.
Values are URL-encoded.
The caller's hashtable is not changed.

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxQueryPages
The maximum number of result pages to retrieve for GET requests.
Default 3.

```yaml
Type: Int16
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: 3
Accept pipeline input: False
Accept wildcard characters: False
```

### -All
Follow the pagination links until every result has been retrieved (ignores -MaxQueryPages).

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

### -Search
A page title to search for with CQL.
Wildcards (* ?) produce a "title ~" search, otherwise an
exact "title =" search.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
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

### System.Management.Automation.PSCustomObject with Results and MultiPage properties.
## NOTES

## RELATED LINKS
