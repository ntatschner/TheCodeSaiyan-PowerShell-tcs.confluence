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

### PageTitle (Default)
```
New-ConfluenceContentInternalLink -PageTitle <String> [-SpaceKey <String>] [-HeadingLink <String>]
 [-LinkText <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### InternalLinkURL
```
New-ConfluenceContentInternalLink [-PageTitle <String>] -InternalLinkURL <String> [-HeadingLink <String>]
 [-LinkText <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### InternalLinkPageId
```
New-ConfluenceContentInternalLink -PageId <String> [-AsUrl] [-HeadingLink <String>] [-LinkText <String>]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
By default New-ConfluenceContentInternalLink returns a Confluence storage-format page link:

    \<ac:link\>\<ri:page ri:content-title="Title" ri:space-key="KEY"/\>\<ac:link-body\>Text\</ac:link-body\>\</ac:link\>

Confluence resolves this link by page title, keeps it working when the page is moved or
renamed, and shows it as a page link.
Give the page with -PageTitle (and -SpaceKey for a page
in another space), or with -PageId, in which case the title and space key are looked up with
Get-ConfluencePage (this needs Set-ConfluenceContext).
-HeadingLink adds ac:anchor so the link
points to a heading on the page.

URL mode (the behaviour before 0.2.0): with -InternalLinkURL, or with -PageId and -AsUrl, an
\<a href\> element pointing to the page URL is returned instead.
With -HeadingLink the URL gets
the Confluence anchor "#\<PageTitle\>-\<Heading\>" with spaces removed; the page title comes from
-PageTitle or is looked up from the page ID in the URL.

All values are escaped, so the result is always well-formed storage format.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentInternalLink -PageTitle 'Runbook' -SpaceKey OPS -HeadingLink 'Restart steps' -LinkText 'Restart steps'
```

Returns \<ac:link ac:anchor="Restart steps"\>\<ri:page ri:content-title="Runbook" ri:space-key="OPS"/\>\<ac:link-body\>Restart steps\</ac:link-body\>\</ac:link\>.

### EXAMPLE 2
```
New-ConfluenceContentInternalLink -PageId 123456 -LinkText 'See the runbook'
```

Looks up page 123456 and returns a storage-format link to it.

### EXAMPLE 3
```
New-ConfluenceContentInternalLink -InternalLinkURL 'https://contoso.atlassian.net/wiki/spaces/DOCS/pages/123/Runbook' -PageTitle 'Runbook' -HeadingLink 'Restart steps'
```

Returns an \<a href\> link to the "Restart steps" heading of the Runbook page (URL mode).

## PARAMETERS

### -PageTitle
The title of the page to link to.
With -InternalLinkURL it is only used to build the heading
anchor.

```yaml
Type: String
Parameter Sets: PageTitle
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: String
Parameter Sets: InternalLinkURL
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpaceKey
The key of the space that contains the page.
Leave it out to link to a page in the same space
as the page that contains the link.

```yaml
Type: String
Parameter Sets: PageTitle
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InternalLinkURL
The full URL of the page to link to (URL mode).

```yaml
Type: String
Parameter Sets: InternalLinkURL
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageId
The ID of the page to link to.
Its title and space key (or, with -AsUrl, its URL) are looked
up.

```yaml
Type: String
Parameter Sets: InternalLinkPageId
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsUrl
With -PageId, return an \<a href\> link to the page URL instead of a storage-format page link.

```yaml
Type: SwitchParameter
Parameter Sets: InternalLinkPageId
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeadingLink
The heading on the target page to link to.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LinkText
The text shown for the link.
Defaults to the page title (URL mode: the URL).

```yaml
Type: String
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

### System.String
## NOTES

## RELATED LINKS
