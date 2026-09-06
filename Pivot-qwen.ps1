
function Pivot-Data {

    [CmdletBinding()]
    param(

        # One or more columns that stay as grouping keys
        [Parameter(Mandatory = $true)]
        [string[]]$GroupBy,

        # Column whose unique values become the new (pivoted) column eaders
        [Parameter(Mandatory = $true)]
        [string]$Pivot,

        # Column whose values are aggregated into the new headers
        [Parameter(Mandatory = $true)]
        [string]$Value,

        # Aggregation to apply to $Value for each group × pivot-value cell
        [Parameter()]
        [ValidateSet('Sum', 'Count', 'Max', 'Min', 'Average')]
        [string]$Operation = 'Sum'
    )

    # ------------------------------------------------------------------
    # PROCESS:  accumulate every pipeline object
    # ------------------------------------------------------------------
    process {
        # $PSItem (aka $_) is the incoming pipeline object
        $collected += $PSItem
    }

    # ------------------------------------------------------------------
    # END:  perform the pivot on the accumulated data
    # ------------------------------------------------------------------
    end {

        if (-not $collected) {
            Write-Warning "Pivot-Data : no input objects were received."
            return
        }

        # ---- 1. Collect unique pivot values in first-seen order ---------
        #      (NOT alphabetically sorted – these are the "pivot column
        #       names" that must keep their natural / insertion order)
        [string[]]$pivotValues = @()
        $seen = @{}
        foreach ($row in $collected) {
            $pv = [string]$row.($Pivot)
            if (-not $seen.ContainsKey($pv)) {
                $seen[$pv] = $true
                $pivotValues += $pv
            }
        }

        # ---- 2. Group all rows by the GroupBy column(s) ------------------
        $groups = $collected | Group-Object -Property $GroupBy

        # ---- 3. Build one output row per group --------------------------
        foreach ($group in $groups) {

            # Use [ordered] so property (column) order is deterministic
            $props = [ordered]@{}

            # -- 3a. Static columns: everything that is NOT $Pivot and NOT 
$Value
            #     They are sorted ALPHABETICALLY by property name.
            $firstRow   = $group.Group[0]
            $staticProps = $firstRow.PSObject.Properties `
                | Where-Object { $_.Name -ne $Pivot -and $_.Name -ne 
$Value } `
                | Sort-Object -Property Name

            foreach ($sp in $staticProps) {
                $props[$sp.Name] = $firstRow.($sp.Name)
            }

            # -- 3b. Pivot columns: one per unique value, in first-seen 
order
            foreach ($pv in $pivotValues) {

                $matches = $group.Group |
                    Where-Object { [string]$_.($Pivot) -eq $pv }

                switch ($Operation) {

                    'Count' {
                        $cell = @($matches).Count
                    }

                    'Sum' {
                        $cell = 0
                        foreach ($m in $matches) {
                            $cell += [double]$m.($Value)
                        }
                    }

                    'Max' {
                        $cell = $null
                        foreach ($m in $matches) {
                            $v = [double]$m.($Value)
                            if ($null -eq $cell -or $v -gt $cell) { $cell 
= $v }
                        }
                        if ($null -eq $cell) { $cell = 0 }
                    }

                    'Min' {
                        $cell = $null
                        foreach ($m in $matches) {
                            $v = [double]$m.($Value)
                            if ($null -eq $cell -or $v -lt $cell) { $cell = $v }
                        }
                        if ($null -eq $cell) { $cell = 0 }
                    }

                    'Average' {
                        if (@($matches).Count -gt 0) {
                            $cell = 0
                            foreach ($m in $matches) { $cell += [double]$m.($Value) }
                            $cell = $cell / @($matches).Count
                        } else {
                            $cell = 0
                        }
                    }
                }

                # Pivot column headers keep their natural (first-seen) 
order
                $props[$pv] = $cell
            }

            # ---- 4. Emit a single PSCustomObject per group ---------------
            [PSCustomObject]$props
        }
    }
}

# =====================================================================================================================================================
# Demo  –  run this block to see it in action
# =====================================================================================================================================================
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

# Single group-by, Sum aggregation
$sample | Pivot-Data -GroupBy 'Month' -Pivot 'Month' -Value 'Sales' -Operation Sum

# Multi group-by, Count aggregation
#$sample | Pivot-Data -GroupBy @('Region','Category') -Pivot 'Month' -Value 'Sales' -Operation Count
