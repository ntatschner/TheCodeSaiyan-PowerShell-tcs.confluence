function ConvertTo-ConfluenceUtcDate {
    <#
    .SYNOPSIS
        Converts a Confluence timestamp (string or DateTime) to a UTC DateTime; DateTime.MinValue if empty or invalid.

    .DESCRIPTION
        PowerShell 7 turns ISO 8601 strings into DateTime values when parsing JSON, Windows
        PowerShell 5.1 leaves them as strings; both are handled without depending on the culture.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param (
        $Value
    )

    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $date = [datetime]::MinValue
    if ($Value -and [datetime]::TryParse("$Value", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$date)) {
        return $date
    }
    return [datetime]::MinValue
}
