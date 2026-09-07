function Add-TotalRow {
    <#
    .SYNOPSIS
        Appends a total row to a stream of PSCustomObjects.
    .DESCRIPTION
        Numeric columns are summed, boolean columns become the percentage of $true
        values, and everything else is left blank. The first column of the total
        row holds the label (default "Total").
    .EXAMPLE
        $data | Add-TotalRow | Format-Table
    .EXAMPLE
        $data | Add-TotalRow -Label 'TOTAL' -Decimals 0
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [PSObject]$InputObject,

        [string]$Label = 'Total',

        [ValidateRange(0, 10)]
        [int]$Decimals = 1
    )

    begin {
        $rows = New-Object System.Collections.Generic.List[object]
        $numericTypes = @([int], [long], [double], [decimal], [single], [int16], [byte], [uint16], [uint32], [uint64])
    }

    process {
        $rows.Add($InputObject)
    }

    end {
        if ($rows.Count -eq 0) { return }

        # Pass the original rows through unchanged
        $rows

        $columns = @($rows[0].PSObject.Properties.Name)
        $total   = [ordered]@{}

        for ($i = 0; $i -lt $columns.Count; $i++) {
            $col = $columns[$i]

            if ($i -eq 0) {
                $total[$col] = $Label
                continue
            }

            $values = @($rows | ForEach-Object { $_.$col } | Where-Object { $null -ne $_ })

            if ($values.Count -eq 0) {
                $total[$col] = $null
            }
            elseif (@($values | Where-Object { $_ -isnot [bool] }).Count -eq 0) {
                # All booleans -> percentage of $true
                $trueCount = @($values | Where-Object { $_ }).Count
                $pct = ($trueCount / $values.Count) * 100
                $total[$col] = ('{0:N' + $Decimals + '}%') -f $pct
            }
            elseif (@($values | Where-Object { $numericTypes -notcontains $_.GetType() }).Count -eq 0) {
                # All numeric -> sum
                $total[$col] = ($values | Measure-Object -Sum).Sum
            }
            else {
                $total[$col] = $null
            }
        }

        [PSCustomObject]$total
    }
}

<# Example
$data = @(
    [PSCustomObject]@{ Name = 'Alice'; Sales = 120; Units = 3; Active = $true  }
    [PSCustomObject]@{ Name = 'Bob';   Sales = 80;  Units = 2; Active = $false }
    [PSCustomObject]@{ Name = 'Cara';  Sales = 200; Units = 5; Active = $true  }
)

$data | Add-TotalRow | Format-Table -AutoSize

# Name  Sales Units Active
# ----  ----- ----- ------
# Alice   120     3   True
# Bob      80     2  False
# Cara    200     5   True
# Total   400    10  66.7%
#>