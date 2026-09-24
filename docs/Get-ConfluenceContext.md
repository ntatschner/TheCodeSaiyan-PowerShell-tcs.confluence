---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-ConfluenceContext

## SYNOPSIS
Returns the Confluence connection set by Set-ConfluenceContext, without the credential.

## SYNTAX

```
Get-ConfluenceContext [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
Get-ConfluenceContext returns the base URL, API endpoint, API version and user name stored by
Set-ConfluenceContext for the current session.
The token is never returned.
Returns nothing
when no context has been set.

This replaces reading the $global:ConfluenceContext variable used by tcs.confluence 0.0.x,
which exposed the Authorization header to every script in the session.

## EXAMPLES

### EXAMPLE 1
```
Get-ConfluenceContext
```

Shows the current connection, for example ConnectionBaseURL https://contoso.atlassian.net.

### EXAMPLE 2
```
if (-not (Get-ConfluenceContext)) { Set-ConfluenceContext -ConfluenceUrl $url -Credential (Get-Credential) }
```

Sets a context only when none exists yet.

## PARAMETERS

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
