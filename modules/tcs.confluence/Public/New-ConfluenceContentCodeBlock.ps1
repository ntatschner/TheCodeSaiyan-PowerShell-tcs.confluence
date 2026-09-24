function New-ConfluenceContentCodeBlock {
    <#
    .SYNOPSIS
        Creates a Confluence code block macro.

    .DESCRIPTION
        New-ConfluenceContentCodeBlock returns the storage-format markup for the Confluence "code"
        macro. The code is placed in a CDATA section, so it is shown exactly as given; a "]]>"
        sequence in the code is split so it cannot end the CDATA section early.

    .PARAMETER Content
        The code to show.

    .PARAMETER Language
        The language used for syntax highlighting, for example powershell, bash or json. Default none.

    .PARAMETER Theme
        The macro theme: Default, Midnight, Eclipse or Emacs.

    .PARAMETER LineNumbers
        Show line numbers.

    .PARAMETER Collapse
        Collapse the code block when the page loads.

    .EXAMPLE
        New-ConfluenceContentCodeBlock -Content 'Get-Process' -Language powershell -LineNumbers

        Returns a PowerShell code block macro with line numbers.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory = $true, HelpMessage = 'The code block content.')]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(HelpMessage = 'The language of the code block.')]
        [string]$Language = 'none',

        [Parameter(HelpMessage = 'The theme of the code block.')]
        [ValidateSet('Default', 'Midnight', 'Eclipse', 'Emacs')]
        [string]$Theme = 'Default',

        [Parameter(HelpMessage = 'Whether to show line numbers.')]
        [switch]$LineNumbers,

        [Parameter(HelpMessage = 'Whether to collapse the code block.')]
        [bool]$Collapse = $false
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
        $LineNumbersValue = if ($LineNumbers) { 'true' } else { 'false' }
        $CollapseValue = if ($Collapse) { 'true' } else { 'false' }
        # "]]>" would end the CDATA section; split it across two sections
        $SafeContent = $Content.Replace(']]>', ']]]]><![CDATA[>')

        $CodeBlockHtml = "<ac:structured-macro ac:name='code'>"
        $CodeBlockHtml += "<ac:parameter ac:name='theme'>$Theme</ac:parameter>"
        $CodeBlockHtml += "<ac:parameter ac:name='linenumbers'>$LineNumbersValue</ac:parameter>"
        $CodeBlockHtml += "<ac:parameter ac:name='collapse'>$CollapseValue</ac:parameter>"
        $CodeBlockHtml += "<ac:parameter ac:name='language'>$([System.Net.WebUtility]::HtmlEncode($Language))</ac:parameter>"
        $CodeBlockHtml += "<ac:plain-text-body><![CDATA[$SafeContent]]></ac:plain-text-body>"
        $CodeBlockHtml += '</ac:structured-macro>'

        return $CodeBlockHtml
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
