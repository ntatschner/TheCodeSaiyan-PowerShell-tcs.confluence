function New-ConfluencePageLayout {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory, HelpMessage = "The name of the layout to create.")]
        [ValidateSet("single", "two_equal", "two_left_sidebar", "two_right_sidebar", "three_equal", "three_with_sidebars")]
        $LayoutType
    )

    DynamicParam {
        $paramDictionary = New-Object -TypeName System.Management.Automation.RuntimeDefinedParameterDictionary
        $variableParams = @(
            @{Name = "SectionOne"; ParameterType = [string]; Mandatory = $true; Position = 1;  HelpMessage = "The first section in the page layout" }
            @{Name = "SectionTwo"; ParameterType = [string]; Mandatory = $true; Position = 2;  HelpMessage = "The second section in the page layout" }
            @{Name = "SectionThree"; ParameterType = [string]; Mandatory = $true; Position = 3;  HelpMessage = "The third section in the page layout" }
        )
        switch -Exact ($LayoutType) {
            "single" {
                $dynamicParamRegistry = @(
                    $variableParams[0]
                )
            }
            "two_equal" {
                $dynamicParamRegistry = @(
                    $variableParams[0],
                    $variableParams[1]
                )
            }
            "two_left_sidebar" {
                $dynamicParamRegistry = @(
                    $variableParams[0],
                    $variableParams[1]
                )
            }
            "two_right_sidebar" {
                $dynamicParamRegistry = @(
                    $variableParams[0],
                    $variableParams[1]
                )
            }
            "three_equal" {
                $dynamicParamRegistry = @(
                    $variableParams[0],
                    $variableParams[1],
                    $variableParams[2]
                )
            }
            "three_with_sidebars" {
                $dynamicParamRegistry = @(
                    $variableParams[0],
                    $variableParams[1],
                    $variableParams[2]
                )
            }
            default {
                return
            }
        }

        foreach ($d in $dynamicParamRegistry) {
            $param = $(New-DynamicParameter @d)
            if (-not $paramDictionary.ContainsKey($param.Name)) {
                $paramDictionary.Add($param.Name, $param.Parameter)
            }
        }

        return $paramDictionary
    }

    process {
        $layoutXml = "<ac:layout>`n"
        $contentSections = 0

        switch ($LayoutType) {
            "single" {
                $layoutXml += @"
  <ac:layout-section ac:type="single">
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionOne"])
     </ac:layout-cell>
  </ac:layout-section>
"@
                $contentSections = 1
            }
            "two_equal" {
                $layoutXml += @"
  <ac:layout-section ac:type="two_equal">
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionOne"])
     </ac:layout-cell>
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionTwo"])
     </ac:layout-cell>
  </ac:layout-section>
"@
                $contentSections = 2
            }
            "two_left_sidebar" {
                $layoutXml += @"
  <ac:layout-section ac:type="two_left_sidebar">
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionOne"])
     </ac:layout-cell>
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionTwo"])
     </ac:layout-cell>
  </ac:layout-section>
"@
                $contentSections = 2
            }
            "two_right_sidebar" {
                $layoutXml += @"
  <ac:layout-section ac:type="two_right_sidebar">
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionOne"])
     </ac:layout-cell>
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionTwo"])
     </ac:layout-cell>
  </ac:layout-section>
"@
                $contentSections = 2
            }
            "three_equal" {
                $layoutXml += @"
  <ac:layout-section ac:type="three_equal">
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionOne"])
     </ac:layout-cell>
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionTwo"])
     </ac:layout-cell>
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionThree"])
     </ac:layout-cell>
  </ac:layout-section>
"@
                $contentSections = 3
            }
            "three_with_sidebars" {
                $layoutXml += @"
  <ac:layout-section ac:type="three_with_sidebars">
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionOne"])
     </ac:layout-cell>
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionTwo"])
     </ac:layout-cell>
     <ac:layout-cell>
        $($PSCmdlet.MyInvocation.BoundParameters["SectionThree"])
     </ac:layout-cell>
  </ac:layout-section>
"@
                $contentSections = 3
            }
        }

        $layoutXml += "`n</ac:layout>"

        return [pscustomobject]@{
            LayoutType = $LayoutType
            LayoutXml = $layoutXml
            ContentSections = $contentSections
        }
    }
}
