---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# Get-ConfluencePageLabel

## SYNOPSIS
Gets the labels of a Confluence page.

## SYNTAX

```
Get-ConfluencePageLabel [-PageId] <String> [[-Prefix] <String>] [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

## DESCRIPTION
Get-ConfluencePageLabel returns the labels of a page through the Confluence v2 API
(GET /wiki/api/v2/pages/\<id\>/labels), reading every result page.
Each label has id, name and
prefix (global, my or team).

## EXAMPLES

### EXAMPLE 1
```
Get-ConfluencePageLabel -PageId 123456 | Select-Object -ExpandProperty name
```

Lists the label names of page 123456.

## PARAMETERS

### -PageId
The ID of the page.
Accepts pipeline input by property name (id), for example from
Get-ConfluencePage.

```yaml
Type: String
Parameter Sets: (All)
Aliases: id

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Prefix
Only return labels with this prefix: global, my or team.

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

### System.Management.Automation.PSCustomObject. One object per label.
## NOTES

## RELATED LINKS
