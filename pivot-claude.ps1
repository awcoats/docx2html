function ConvertTo-TransposedObject {
    <#
    .SYNOPSIS
        Transposes objects from the pipeline: properties become rows, objects become columns.

    .PARAMETER InputObject
        Objects to transpose (PSCustomObjects, Import-Csv output, Get-Process, etc.).

    .PARAMETER KeyProperty
        Optional. Property whose values become the new column headers.
        If omitted, columns are named Row1, Row2, ...

    .PARAMETER LabelColumnName
        Name of the first output column, which holds the original property names.
        Defaults to 'Property'.

    .EXAMPLE
        Import-Csv .\servers.csv | ConvertTo-TransposedObject -KeyProperty Name

    .EXAMPLE
        Get-Process | Select-Object -First 3 Name, Id, WS |
            ConvertTo-TransposedObject -KeyProperty Name | Format-Table -AutoSize
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject[]]$InputObject,

        [string]$KeyProperty,

        [string]$LabelColumnName = 'Property'
    )

    begin {
        $rows = New-Object System.Collections.ArrayList
    }

    process {
        foreach ($item in $InputObject) {
            [void]$rows.Add($item)
        }
    }

    end {
        if ($rows.Count -eq 0) { return }

        # Build unique column headers for the output.
        $headers = @()
        $seen    = @{}
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $name = if ($KeyProperty) { [string]$rows[$i].$KeyProperty } else { "Row$($i + 1)" }
            if ([string]::IsNullOrEmpty($name)) { $name = "Row$($i + 1)" }

            # De-duplicate headers so the ordered hashtable never collides.
            $base = $name; $n = 2
            while ($seen.ContainsKey($name)) { $name = "$base($n)"; $n++ }
            $seen[$name] = $true
            $headers += $name
        }

        # Collect property names across all rows, preserving first-seen order.
        $propNames = @()
        foreach ($row in $rows) {
            foreach ($p in $row.PSObject.Properties.Name) {
                if ($p -ne $KeyProperty -and $propNames -notcontains $p) {
                    $propNames += $p
                }
            }
        }

        # Emit one output object per original property.
        foreach ($prop in $propNames) {
            $out = [ordered]@{ $LabelColumnName = $prop }
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $out[$headers[$i]] = $rows[$i].$prop
            }
            [PSCustomObject]$out
        }
    }
}

$sample = @(
    [PSCustomObject]@{ Region = 'North'; Month = 'Jan'; Category = 'A'; 
Sales = 100 }
    [PSCustomObject]@{ Region = 'North'; Month = 'Feb'; Category = 'B'; 
Sales = 200 }
    [PSCustomObject]@{ Region = 'South'; Month = 'Jan'; Category = 'A'; 
Sales = 150 }
    [PSCustomObject]@{ Region = 'South'; Month = 'Feb'; Category = 'B'; 
Sales = 250 }
    [PSCustomObject]@{ Region = 'North'; Month = 'Mar'; Category = 'A'; 
Sales =  80 }
    [PSCustomObject]@{ Region = 'South'; Month = 'Mar'; Category = 'B'; 
Sales = 300 }
)

$sample | ConvertTo-TransposedObject -KeyProperty Month -LabelColumnName 'Property' | Format-Table -AutoSize