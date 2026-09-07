<#
.SYNOPSIS
    Appends a "Total" row to a pipeline of PSCustomObjects.

.DESCRIPTION
    - Boolean columns  : counts True values and renders the percentage of 
total rows.
    - Numeric columns  : sums all Int, Float (Single), and Double values.
    - All other columns: left blank in the total row.

.EXAMPLE
    $data | Add-TotalRow
#>
function Add-TotalRow {

    [CmdletBinding()]
    param()

    # ---------- collection ----------
    $items = [System.Collections.Generic.List[object]]::new()

    process {
        $items.Add($PSItem)
    }

    end {
        if ($items.Count -eq 0) { return }

        # Union of every property name across every object
        $allProperties = $items |
            ForEach-Object  { $_.PSObject.Properties.Name } |
            Sort-Object -Unique

        $totalValues = @{}

        foreach ($prop in $allProperties) {

            # Grab every non-null value for this column; force an array
            $values = @(
                $items |
                    ForEach-Object { $_.$prop } |
                    Where-Object  { $null -ne $_ }
            )

            # Skip columns that are entirely null
            if ($values.Count -eq 0) {
                $totalValues[$prop] = ''
                continue
            }

            # ----- Boolean column -----
            # True when *every* value in the column is a [bool]
            $nonBool = @($values | Where-Object { $_ -isnot [bool] })

            if ($nonBool.Count -eq 0) {
                $trueCount = @($values | Where-Object { $_ -eq $true 
}).Count
                $pct = [math]::Round(($trueCount / $items.Count * 100), 2)
                $totalValues[$prop] = '{0:N2} %' -f $pct
                continue
            }

            # ----- Numeric column (Int / Float / Double) -----
            $nonNumeric = @(
                $values | Where-Object {
                    $_ -isnot [int]    -and
                    $_ -isnot [float]  -and
                    $_ -isnot [double]
                }
            )

            if ($nonNumeric.Count -eq 0) {
                # Cast every value to [double] so mixed int/float/double 
sum cleanly
                $sum = [double]0
                foreach ($v in $values) {
                    $sum += [double]$v
                }
                # If the result is a whole number, drop the decimals for 
cleanliness
                if ($sum -eq [math]::Floor($sum)) {
                    $totalValues[$prop] = [long]$sum
                } else {
                    $totalValues[$prop] = [double]($sum)
                }
                continue
            }

            # ----- Anything else (string, etc.) -----
            $totalValues[$prop] = ''
        }

        # Build the total row as a PSCustomObject with the same shape
        $totalRow = [pscustomobject]$totalValues

        # Emit original objects followed by the total row
        foreach ($item in $items) {
            $item
        }
        $totalRow
    }
}
