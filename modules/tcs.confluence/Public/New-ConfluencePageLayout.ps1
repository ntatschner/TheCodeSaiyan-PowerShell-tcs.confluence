function New-ConfluencePageLayout {
    <#
    .SYNOPSIS
        Creates a Confluence page layout with one or more sections of one to three columns.

    .DESCRIPTION
        New-ConfluencePageLayout returns the storage-format markup for an ac:layout.

        For a layout with one section, give -LayoutType and the content of each column with
        -SectionOne, -SectionTwo and -SectionThree: single takes one column, two_equal,
        two_left_sidebar and two_right_sidebar take two, three_equal and three_with_sidebars take
        three. Passing the wrong number of columns is an error.

        For a layout with several sections, pass -Section with one hashtable per section, each with a
        LayoutType and a Cells array, for example @{ LayoutType = 'two_equal'; Cells = '<p>a</p>', '<p>b</p>' }.

        The column content is inserted as given; it is expected to be storage-format markup, such as
        the output of the New-ConfluenceContent* functions.

        Before 0.2.0 this function returned an object with LayoutType, LayoutXml and ContentSections;
        it now returns the markup string like the other builders.

    .PARAMETER LayoutType
        The layout of a single section: single, two_equal, two_left_sidebar, two_right_sidebar,
        three_equal or three_with_sidebars.

    .PARAMETER SectionOne
        The content of the first column.

    .PARAMETER SectionTwo
        The content of the second column (two- and three-column layouts).

    .PARAMETER SectionThree
        The content of the third column (three-column layouts).

    .PARAMETER Section
        Several sections, in order: hashtables with LayoutType and Cells (one string per column).

    .EXAMPLE
        New-ConfluencePageLayout -LayoutType two_equal -SectionOne '<p>Left</p>' -SectionTwo '<p>Right</p>'

        Returns a layout with one two-column section.

    .EXAMPLE
        New-ConfluencePageLayout -Section @{ LayoutType = 'single'; Cells = '<h1>Report</h1>' }, @{ LayoutType = 'two_equal'; Cells = '<p>Left</p>', '<p>Right</p>' }

        Returns a layout with a full-width section followed by a two-column section.

    .OUTPUTS
        System.String
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Only builds a storage-format string in memory; nothing outside the session is changed.')]
    [CmdletBinding(DefaultParameterSetName = 'Single')]
    [OutputType([string])]
    param (
        [Parameter(Mandatory, Position = 0, ParameterSetName = 'Single', HelpMessage = 'The name of the layout to create.')]
        [ValidateSet('single', 'two_equal', 'two_left_sidebar', 'two_right_sidebar', 'three_equal', 'three_with_sidebars')]
        [string]$LayoutType,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'Single', HelpMessage = 'The first section in the page layout.')]
        [AllowEmptyString()]
        [string]$SectionOne,

        [Parameter(Position = 2, ParameterSetName = 'Single', HelpMessage = 'The second section in the page layout.')]
        [string]$SectionTwo,

        [Parameter(Position = 3, ParameterSetName = 'Single', HelpMessage = 'The third section in the page layout.')]
        [string]$SectionThree,

        [Parameter(Mandatory, ParameterSetName = 'Multiple', HelpMessage = 'Sections: hashtables with LayoutType and Cells.')]
        [ValidateNotNullOrEmpty()]
        [hashtable[]]$Section
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
        $cellCounts = @{
            single              = 1
            two_equal           = 2
            two_left_sidebar    = 2
            two_right_sidebar   = 2
            three_equal         = 3
            three_with_sidebars = 3
        }

        if ($PSCmdlet.ParameterSetName -eq 'Single') {
            $cells = @($SectionOne)
            if ($PSBoundParameters.ContainsKey('SectionTwo')) { $cells += $SectionTwo }
            if ($PSBoundParameters.ContainsKey('SectionThree')) {
                if (-not $PSBoundParameters.ContainsKey('SectionTwo')) {
                    throw 'SectionThree needs SectionTwo.'
                }
                $cells += $SectionThree
            }
            $Section = @(@{ LayoutType = $LayoutType; Cells = $cells })
        }

        $markup = '<ac:layout>'
        $number = 0
        foreach ($item in $Section) {
            $number++
            $type = [string]$item['LayoutType']
            if (-not $cellCounts.ContainsKey($type)) {
                throw "Section $number has an unknown LayoutType '$type'. Valid values: $(($cellCounts.Keys | Sort-Object) -join ', ')."
            }
            $sectionCells = @($item['Cells'])
            if ($sectionCells.Count -ne $cellCounts[$type]) {
                throw "Section $number ($type) needs $($cellCounts[$type]) column(s) but $($sectionCells.Count) were given."
            }
            $markup += "<ac:layout-section ac:type=`"$type`">"
            foreach ($cell in $sectionCells) {
                $markup += "<ac:layout-cell>$cell</ac:layout-cell>"
            }
            $markup += '</ac:layout-section>'
        }
        $markup += '</ac:layout>'
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
