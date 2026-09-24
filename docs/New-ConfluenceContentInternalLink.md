---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentInternalLink

## SYNOPSIS
Creates a link to a Confluence page, optionally to a heading on that page.

## SYNTAX

### InternalLinkURL (Default)
```
New-ConfluenceContentInternalLink -InternalLinkURL <String> [-PageTitle <String>] [-HeadingLink <String>]
 [-LinkText <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### InternalLinkPageId
```
New-ConfluenceContentInternalLink -PageId <String> [-HeadingLink <String>] [-LinkText <String>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentInternalLink returns an \<a\> element that links to a Confluence page given
by its URL or by its page ID.
With -PageId the page URL is looked up with Get-ConfluencePage,
which needs a Confluence context (Set-ConfluenceContext).

With -HeadingLink the link points to a heading on the page, using the Confluence anchor format
"#\<PageTitle\>-\<Heading\>" with spaces removed.
The page title comes from -PageTitle or, when not
given, is looked up from the page ID in the URL.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentInternalLink -InternalLinkURL 'https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook' -PageTitle 'Runbook' -HeadingLink 'Restart steps' -LinkText 'Restart steps'
```

Returns a link to the "Restart steps" heading of the Runbook page.

### EXAMPLE 2
```
New-ConfluenceContentInternalLink -PageId 123456 -LinkText 'See the runbook'
```

Looks up page 123456 and returns a link to it.

## PARAMETERS

### -InternalLinkURL
The full URL of the page to link to.

```yaml
Type:String
Parameter Sets: InternalLinkURL
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

### -PageTitle
The title of the page, used to build the heading anchor.

```yaml
Type:String
Parameter Sets: InternalLinkURL
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

### -PageId
The ID of the page to link to.

```yaml
Type:String
Parameter Sets: InternalLinkPageId
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

### -HeadingLink
The heading on the target page to link to.

```yaml
Type:String
Parameter Sets:   (All)
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

### -LinkText
The text shown for the link.
Defaults to the URL.

```yaml
Type:String
Parameter Sets:   (All)
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

### System.String
## NOTES

## RELATED LINKS
