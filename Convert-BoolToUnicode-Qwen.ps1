function Convert-BoolToUnicode {
    <#
    .SYNOPSIS
        Converts boolean values (or boolean properties of objects) to
        unicode glyphs: a green check mark for $true, a ballot X for $false.

    .DESCRIPTION
        Pipeline-aware function for PowerShell 5.1. Accepts raw [bool]
        values or objects (PSCustomObject, Hashtable, etc.) and replaces
        every boolean member with the corresponding unicode character.

    .PARAMETER Value
        A single boolean, or an object whose boolean properties should be
        converted.  Accepts pipeline input.

    .PARAMETER CheckMark
        Unicode char for $true.  Default: U+2705 (✅, green check mark).

    .PARAMETER CrossMark
        Unicode char for $false.  Default: U+2717 (✗, ballot X).

    .PARAMETER Reformat
        When set, the function re-outputs the original object (or hashtable)
        with its boolean properties swapped out.  Without it, the function
        emits one string per pipeline item.

    .EXAMPLE
        $true, $false, $true | Convert-BoolToUnicode
        ✅
        ✗
        ✅

    .EXAMPLE
        [pscustomobject]@{ Name = 'Server'; Online = $true;  DiskFull = $false } |
        Convert-BoolToUnicode -Reformat
        Name   Online  DiskFull
        ----   ------  --------
        Server  ✅     ✗
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]$Value,

        [string]$CheckMark = [char]0x2705,   # ✅  green check mark
        [string]$CrossMark = [char]0x2717    # ✗  ballot X
    )

    begin {
        # Pre-compute the two glyphs once so the process block stays hot.
        $script:_glyphTrue  = $CheckMark
        $script:_glyphFalse = $CrossMark
    }

    process {
        # ---- 1.  The item itself IS a [bool] -------------------------
        if ($Value -is [bool]) {
            if ($Value) { $script:_glyphTrue } else { $script:_glyphFalse }
            return
        }

        # ---- 2.  The item is a Hashtable -----------------------------
        if ($Value -is [hashtable]) {
            $newH = @{}
            foreach ($key in $Value.Keys) {
                $v = $Value[$key]
                if ($v -is [bool]) {
                    $newH[$key] = if ($v) { $script:_glyphTrue } else { $script:_glyphFalse }
                } else {
                    $newH[$key] = $v
                }
            }
            Write-Output $newH
            return
        }

        # ---- 3.  The item is a PSCustomObject / .NET object ----------
        if ($Value -is [psobject] -or $Value -is [object]) {
            $props = $Value.PSObject.Properties
            $newObj = [pscustomobject]@{}
            $changed = $false

            foreach ($prop in $props) {
                if ($prop.Value -is [bool]) {
                    $val = if ($prop.Value) { $script:_glyphTrue } else { $script:_glyphFalse }
                    $newObj | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $val -Force
                    $changed = $true
                } else {
                    $newObj | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -Force
                }
            }

            if ($changed) {
                Write-Output $newObj
            } else {
                # No bools found – pass the original through unchanged.
                Write-Output $Value
            }
            return
        }

        # ---- 4.  Fallback: not a bool, not an object we recognise ----
        Write-Output $Value
    }
}