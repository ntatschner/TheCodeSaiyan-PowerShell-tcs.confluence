function Get-ConfluenceSpaces {
    [CmdletBinding(DefaultParameterSetName = "AllSpaces")]
    param (
        [Parameter(ParameterSetName = "SpaceId")]
        [string]$SpaceId,
        
        [Parameter(ParameterSetName = "AllSpaces", HelpMessage = "Name filter (supports wildcard * ?) applied client-side).")]
        [string]$Search,
        
        [int]$ResultsLimit = 50,
        [int16]$MaxQueryPages = 3
    )

    process {
        try {
            if ($PSCmdlet.ParameterSetName -eq 'SpaceId') {
                $resp = Invoke-ConfluenceRequest -Method GET -Resource spaces -Id $SpaceId -MaxQueryPages 1 -ErrorAction Stop
                return $resp
            } else {
                $query = @{ limit = $ResultsLimit }
                $resp  = Invoke-ConfluenceRequest -Method GET -Resource spaces -Query $query -MaxQueryPages $MaxQueryPages -ErrorAction Stop
                $results = $resp.Results
                if ($Search) {
                    # Simple case-insensitive like; allow wildcard * ? translation to -like
                    $pattern = $Search
                    $results = $results | Where-Object { $_.name -like $pattern }
                }
                return $results
            }
        } catch {
            Write-Error "Failed to retrieve space information. Error: $_"
        }
    }
}
