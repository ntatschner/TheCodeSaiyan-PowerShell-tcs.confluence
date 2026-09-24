function New-ConfluenceContentJiraIssue {
    <#
    .SYNOPSIS
        Creates a Confluence Jira macro for one issue or a JQL query.

    .DESCRIPTION
        New-ConfluenceContentJiraIssue returns the storage-format markup for the Confluence "jira"
        macro. With -IssueKey it shows a single issue (key, summary and status); with -JqlQuery it
        shows a table of the issues the query returns.

        On Confluence Cloud connected to one Jira site the macro works without -Server and -ServerId.
        When several Jira sites are linked, pass the application link name (-Server) and ID
        (-ServerId); the easiest way to find them is to insert the macro once in the editor and look
        at the page's storage format.

    .PARAMETER IssueKey
        The key of the issue, for example OPS-123.

    .PARAMETER JqlQuery
        A JQL query, for example 'project = OPS AND status = Open'.

    .PARAMETER Columns
        With -JqlQuery, the columns to show, for example key, summary, status, assignee.

    .PARAMETER MaximumIssues
        With -JqlQuery, the maximum number of issues to show. Default 20.

    .PARAMETER ShowSummary
        With -IssueKey, whether to show the issue summary next to the key. Default $true.

    .PARAMETER Server
        The name of the Jira application link.

    .PARAMETER ServerId
        The ID of the Jira application link.

    .EXAMPLE
        New-ConfluenceContentJiraIssue -IssueKey OPS-123

        Returns a macro that shows issue OPS-123.

    .EXAMPLE
        New-ConfluenceContentJiraIssue -JqlQuery 'project = OPS AND resolution = Unresolved' -Columns key, summary, status

        Returns a macro that shows the open OPS issues in a table.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding(DefaultParameterSetName = 'Issue')]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, ParameterSetName = 'Issue', HelpMessage = 'The issue key, for example OPS-123.')]
        [ValidatePattern('^[A-Za-z][A-Za-z0-9_]*-\d+$')]
        [string]$IssueKey,

        [Parameter(Mandatory, ParameterSetName = 'Jql', HelpMessage = 'The JQL query.')]
        [ValidateNotNullOrEmpty()]
        [string]$JqlQuery,

        [Parameter(ParameterSetName = 'Jql', HelpMessage = 'The columns to show.')]
        [string[]]$Columns,

        [Parameter(ParameterSetName = 'Jql', HelpMessage = 'The maximum number of issues to show.')]
        [ValidateRange(1, 1000)]
        [int]$MaximumIssues = 20,

        [Parameter(ParameterSetName = 'Issue', HelpMessage = 'Show the issue summary.')]
        [bool]$ShowSummary = $true,

        [Parameter(HelpMessage = 'The Jira application link name.')]
        [string]$Server,

        [Parameter(HelpMessage = 'The Jira application link ID.')]
        [string]$ServerId
    )

    $TelemetryArgs = @{
        ModuleName    = $MyInvocation.MyCommand.Module.Name
        ModuleVersion = [string]$MyInvocation.MyCommand.Module.Version
        CommandName   = $MyInvocation.MyCommand.Name
        ExecutionID   = [guid]::NewGuid().ToString()
    }
    Invoke-TelemetryCollection @TelemetryArgs -Stage Start -ClearTimer
    $telemetryFailed = $false
    try {
        $parameters = New-Object -TypeName System.Collections.Generic.List[string]
        if ($PSCmdlet.ParameterSetName -eq 'Issue') {
            $parameters.Add('<ac:parameter ac:name="key">' + (ConvertTo-ConfluenceXmlText -Text $IssueKey.ToUpperInvariant()) + '</ac:parameter>')
            $showSummaryValue = if ($ShowSummary) { 'true' } else { 'false' }
            $parameters.Add("<ac:parameter ac:name=`"showSummary`">$showSummaryValue</ac:parameter>")
        }
        else {
            $parameters.Add('<ac:parameter ac:name="jqlQuery">' + (ConvertTo-ConfluenceXmlText -Text $JqlQuery) + '</ac:parameter>')
            if ($Columns) {
                $parameters.Add('<ac:parameter ac:name="columns">' + (ConvertTo-ConfluenceXmlText -Text ($Columns -join ',')) + '</ac:parameter>')
            }
            $parameters.Add("<ac:parameter ac:name=`"maximumIssues`">$MaximumIssues</ac:parameter>")
        }
        if ($Server) { $parameters.Add('<ac:parameter ac:name="server">' + (ConvertTo-ConfluenceXmlText -Text $Server) + '</ac:parameter>') }
        if ($ServerId) { $parameters.Add('<ac:parameter ac:name="serverId">' + (ConvertTo-ConfluenceXmlText -Text $ServerId) + '</ac:parameter>') }
        return ('<ac:structured-macro ac:name="jira">' + ($parameters -join '') + '</ac:structured-macro>')
    }
    catch {
        $telemetryFailed = $true
        Invoke-TelemetryCollection @TelemetryArgs -Stage End -Failed $true -Exception $_
        throw
    }
    finally {
        if (-not $telemetryFailed) {
            Invoke-TelemetryCollection @TelemetryArgs -Stage End
        }
    }
}
