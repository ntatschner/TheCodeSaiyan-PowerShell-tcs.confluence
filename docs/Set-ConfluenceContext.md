---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Set-ConfluenceContext

## SYNOPSIS
Sets the Confluence site and credential used by the other tcs.confluence commands.

## SYNTAX

### Token (Default)
```
Set-ConfluenceContext -ConfluenceUrl <String> -Username <String> -PersonalAccessToken <String>
 [-ApiVersion <String>] [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Credential
```
Set-ConfluenceContext -ConfluenceUrl <String> -Credential <PSCredential> [-ApiVersion <String>]
 [-ProgressAction <ActionPreference>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Set-ConfluenceContext stores the Confluence base URL and your credential for the current
PowerShell session.
The URL is normalised (a trailing /wiki, /wiki/api/v2 or /wiki/rest/api is
removed) and the API endpoint for the chosen version is derived from it.

The credential is kept in memory only, as a PSCredential (the token is held in a SecureString).
It is never written to disk, never returned by Get-ConfluenceContext and never written to the
verbose, warning or error streams.
The Authorization header is built for each request.

Use Get-ConfluenceContext to see the current connection details.

## EXAMPLES

### EXAMPLE 1
```
Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net' -Credential (Get-Credential)
```

Prompts for the e-mail address and API token and stores them for the session.

### EXAMPLE 2
```
Set-ConfluenceContext -ConfluenceUrl 'https://contoso.atlassian.net/wiki' -Username 'me@contoso.com' -PersonalAccessToken $token
```

Uses a token held in a variable.
The URL is normalised to https://contoso.atlassian.net.

## PARAMETERS

### -ConfluenceUrl
The base URL of the Confluence site, for example https://contoso.atlassian.net.
Only https URLs
are accepted so the credential is never sent in clear text.

```yaml
Type:String
Parameter Sets:   (All)
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

### -Username
The account e-mail address (Confluence Cloud) or user name used with the API token.

```yaml
Type:String
Parameter Sets: Token
Aliases:EmailAddress
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

### -PersonalAccessToken
The API token or personal access token as plain text.
Kept for backward compatibility; prefer
-Credential so the token is not visible in your command history.

```yaml
Type:String
Parameter Sets: Token
Aliases:PAT
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

### -Credential
A PSCredential whose user name is the account e-mail address and whose password is the API
token, for example from Get-Credential or a secret store.

```yaml
Type:PSCredential
Parameter Sets: Credential
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

### -ApiVersion
The default REST API version for the derived ConnectionURI: v2 (default) or v1.

```yaml
Type:String
Parameter Sets:   (All)
Aliases:
Required: False
Position:Named
Default value: None
Default value: None
Default value: V2
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

### None.
## NOTES

## RELATED LINKS
