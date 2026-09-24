function New-ConfluenceContentStatus {
    <#
    .SYNOPSIS
        Creates a Confluence status lozenge (status macro).

    .DESCRIPTION
        New-ConfluenceContentStatus returns the storage-format markup for the Confluence "status"
        macro: a small coloured label such as DONE or IN PROGRESS. The text is escaped. Use the result
        in a paragraph or, with New-ConfluenceContentTable -Raw, in a table cell.

    .PARAMETER Text
        The text of the lozenge. Confluence shows it in capitals.

    .PARAMETER Colour
        The colour: Grey (default), Red, Yellow, Green, Blue or Purple. -Color is an alias.

    .PARAMETER Subtle
        Use the subtle (outlined) style instead of a filled lozenge.

    .EXAMPLE
        New-ConfluenceContentStatus -Text 'Done' -Colour Green

        Returns <ac:structured-macro ac:name="status"><ac:parameter ac:name="colour">Green</ac:parameter><ac:parameter ac:name="title">Done</ac:parameter></ac:structured-macro>.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, HelpMessage = 'The text of the status lozenge.')]
        [string]$Text,

        [Parameter(HelpMessage = 'The colour of the status lozenge.')]
        [Alias('Color')]
        [ValidateSet('Grey', 'Red', 'Yellow', 'Green', 'Blue', 'Purple')]
        [string]$Colour = 'Grey',

        [Parameter(HelpMessage = 'Use the subtle (outlined) style.')]
        [switch]$Subtle
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
        # Normalise the casing the macro expects (Grey, Red ...)
        $colourName = $Colour.Substring(0, 1).ToUpperInvariant() + $Colour.Substring(1).ToLowerInvariant()
        $markup = '<ac:structured-macro ac:name="status">'
        $markup += "<ac:parameter ac:name=`"colour`">$colourName</ac:parameter>"
        $markup += '<ac:parameter ac:name="title">' + (ConvertTo-ConfluenceXmlText -Text $Text) + '</ac:parameter>'
        if ($Subtle) { $markup += '<ac:parameter ac:name="subtle">true</ac:parameter>' }
        $markup += '</ac:structured-macro>'
        return $markup
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
