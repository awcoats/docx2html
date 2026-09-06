function ConvertTo-TransposedObject {
    <#
    .SYNOPSIS
        Transposes objects from the pipeline: properties become rows, objects become columns.
        Objects sharing the same key value are merged into one column and their numeric
        values are summed.

    .PARAMETER InputObject
        Objects to transpose (PSCustomObjects, Import-Csv output, Get-Process, etc.).

    .PARAMETER KeyProperty
        Optional. Property whose values become the new column headers.
        If omitted, columns are named Row1, Row2, ...

    .PARAMETER LabelColumnName
        Name of the first output column, which holds the original property names.
        Defaults to 'Property'.

    .EXAMPLE
        Import-Csv .\usage.csv | ConvertTo-TransposedObject -KeyProperty Server

    .EXAMPLE
        Get-Process | Select-Object Name, WS, CPU |
            ConvertTo-TransposedObject -KeyProperty Name | Format-Table -AutoSize
        # All chrome.exe processes collapse into a single 'chrome' column with WS/CPU summed.
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

        # columns: ordered header -> ordered hashtable of property -> aggregated value
        $columns   = [ordered]@{}
        $propNames = @()

        for ($i = 0; $i -lt $rows.Count; $i++) {
            $row = $rows[$i]

            $header = if ($KeyProperty) { [string]$row.$KeyProperty } else { "Row$($i + 1)" }
            if ([string]::IsNullOrEmpty($header)) { $header = "Row$($i + 1)" }

            if (-not $columns.Contains($header)) { $columns[$header] = [ordered]@{} }
            $col = $columns[$header]

            foreach ($p in $row.PSObject.Properties) {
                if ($p.Name -eq $KeyProperty) { continue }
                if ($propNames -notcontains $p.Name) { $propNames += $p.Name }

                $value = $p.Value

                if (-not $col.Contains($p.Name)) {
                    # First occurrence: store numbers as numbers, everything else as-is.
                    $num = 0.0
                    if ($null -ne $value -and [double]::TryParse([string]$value, [ref]$num)) {
                        $col[$p.Name] = $num
                    } else {
                        $col[$p.Name] = $value
                    }
                    continue
                }

                # Duplicate key: sum if both sides are numeric, otherwise keep the first value.
                $existing = $col[$p.Name]
                $a = 0.0; $b = 0.0
                $existingIsNum = ($null -ne $existing) -and [double]::TryParse([string]$existing, [ref]$a)
                $valueIsNum    = ($null -ne $value)    -and [double]::TryParse([string]$value,    [ref]$b)

                if ($existingIsNum -and $valueIsNum) {
                    $col[$p.Name] = $a + $b
                }
                elseif ($null -eq $existing -or [string]::IsNullOrEmpty([string]$existing)) {
                    $col[$p.Name] = $value
                }
            }
        }

        # Emit one output object per original property.
        foreach ($prop in $propNames) {
            $out = [ordered]@{ $LabelColumnName = $prop }
            foreach ($header in $columns.Keys) {
                $out[$header] = $columns[$header][$prop]
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