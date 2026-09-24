---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentJiraIssue

## SYNOPSIS
Creates a Confluence Jira macro for one issue or a JQL query.

## SYNTAX

### Issue (Default)
```
New-ConfluenceContentJiraIssue -IssueKey <String> [-ShowSummary <Boolean>] [-Server <String>]
 [-ServerId <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### Jql
```
New-ConfluenceContentJiraIssue -JqlQuery <String> [-Columns <String[]>] [-MaximumIssues <Int32>]
 [-Server <String>] [-ServerId <String>] [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentJiraIssue returns the storage-format markup for the Confluence "jira"
macro.
With -IssueKey it shows a single issue (key, summary and status); with -JqlQuery it
shows a table of the issues the query returns.

On Confluence Cloud connected to one Jira site the macro works without -Server and -ServerId.
When several Jira sites are linked, pass the application link name (-Server) and ID
(-ServerId); the easiest way to find them is to insert the macro once in the editor and look
at the page's storage format.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentJiraIssue -IssueKey OPS-123
```

Returns a macro that shows issue OPS-123.

### EXAMPLE 2
```
New-ConfluenceContentJiraIssue -JqlQuery 'project = OPS AND resolution = Unresolved' -Columns key, summary, status
```

Returns a macro that shows the open OPS issues in a table.

## PARAMETERS

### -IssueKey
The key of the issue, for example OPS-123.

```yaml
Type: String
Parameter Sets: Issue
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -JqlQuery
A JQL query, for example 'project = OPS AND status = Open'.

```yaml
Type: String
Parameter Sets: Jql
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Columns
With -JqlQuery, the columns to show, for example key, summary, status, assignee.

```yaml
Type: String[]
Parameter Sets: Jql
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumIssues
With -JqlQuery, the maximum number of issues to show.
Default 20.

```yaml
Type: Int32
Parameter Sets: Jql
Aliases:

Required: False
Position: Named
Default value: 20
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowSummary
With -IssueKey, whether to show the issue summary next to the key.
Default $true.

```yaml
Type: Boolean
Parameter Sets: Issue
Aliases:

Required: False
Position: Named
Default value: True
Accept pipeline input: False
Accept wildcard characters: False
```

### -Server
The name of the Jira application link.

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

### -ServerId
The ID of the Jira application link.

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
