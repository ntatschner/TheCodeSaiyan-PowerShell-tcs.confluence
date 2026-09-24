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
 [[-MaxQueryPages] <Int16>] [[-Search] <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
Invoke-ConfluenceRequest is the transport used by every tcs.confluence REST command.
It builds
the endpoint from -Resource (or an explicit -URIPath), appends query parameters and an optional
CQL title search, sends the request with the credential stored by Set-ConfluenceContext and
follows "next" pagination links for GET requests up to -MaxQueryPages pages.

Behaviour that is applied automatically:
- A wildcard (* or ?) -Search against v2 pages switches to the v1 content API, which supports CQL.
- A spaceKey query value for v2 pages is resolved to a spaceId (numeric keys are used as-is).
- Legacy short paths (/pages, /content, /spaces) are mapped to their full API paths.
- Pagination links pointing to a different host are not followed, so the credential is only
  ever sent to the Confluence site in the context.

HTTP error responses are reported as errors that include the status and the Confluence error
titles; use -ErrorAction Stop to turn them into terminating errors.

Returns an object with a Results array (all collected results) and a MultiPage flag.

## EXAMPLES

### EXAMPLE 1
```
Invoke-ConfluenceRequest -Method GET -Resource pages -Query @{ spaceKey = 'DOCS'; limit = 50 }
```

Returns up to three pages of results for the pages in the DOCS space.

### EXAMPLE 2
```
Invoke-ConfluenceRequest -Method GET -Resource pages -Search 'Release*'
```

Searches for pages whose title starts with "Release" using the v1 CQL search.

### EXAMPLE 3
```
Invoke-ConfluenceRequest -Method DELETE -Resource pages -Id 123456 -ErrorAction Stop
```

Deletes page 123456 and throws if Confluence reports an error.

## PARAMETERS

### -Method
The HTTP method: GET, POST, PUT or DELETE.

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

### -URIPath
An explicit API path such as /wiki/api/v2/pages.
Used as-is unless -Resource is given without
-URIPath.

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

### -Resource
A resource shortcut used to build the path: pages, content or spaces.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 3Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -ApiVersion
The API version used with -Resource: 2 (default, /wiki/api/v2) or 1 (/wiki/rest/api).

```yaml
Type:
Int32
Parameter Sets:   (All)
Aliases:
Required: False
Position: 4Default
Default value: None
Default value: 2
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Id
A resource ID appended to the path built from -Resource, for example a page ID.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 5Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -RawPath
Do not build the path from -Resource; use -URIPath exactly as given.

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

### -Body
The JSON request body for POST and PUT requests.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 6Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Query
A hashtable of query-string parameters, for example @{ limit = 25; spaceKey = 'DOCS' }.
Values are URL-encoded.

```yaml
Type:Hashtable
Parameter Sets:   (All)
Aliases:
Required: False
Position: 7Default
Default value: None
Default value: None
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -MaxQueryPages
The maximum number of result pages to retrieve for GET requests.
Default 3.

```yaml
Type:
Int16
Parameter Sets:   (All)
Aliases:
Required: False
Position: 8Default
Default value: None
Default value: 3
Accept pipeline input: False
input:False
Accept pipeline input: False
Accept wildcard characters: False
Accept wildcard characters: False
```

### -Search
A page title to search for.
Wildcards (* ?) produce a CQL "title ~" search, otherwise an exact
"title =" search.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position: 9Default
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

### System.Management.Automation.PSCustomObject with Results and MultiPage properties.
## NOTES

## RELATED LINKS
