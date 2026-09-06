function ConvertTo-CheckMark {
    <#
    .SYNOPSIS
        Converts boolean values to Unicode check mark / X symbols.

    .DESCRIPTION
        Accepts booleans or objects from the pipeline. A raw boolean is
        replaced with a check mark (true) or X (false). For objects, a
        copy is returned with each boolean property replaced by the
        matching symbol; non-boolean and $null values are left unchanged.
        The original input objects are not modified.

    .PARAMETER InputObject
        A boolean or an object with boolean properties. Accepts pipeline
        input. $null is passed through untouched.

    .PARAMETER Property
        One or more property names to convert. When omitted, every
        boolean property on the object is converted. Ignored for raw
        boolean input.

    .PARAMETER TrueMark
        String to use for $true values. Defaults to a green check mark
        emoji (U+2705).

    .PARAMETER FalseMark
        String to use for $false values. Defaults to a red X emoji
        (U+274C).

    .EXAMPLE
        $true, $false | ConvertTo-CheckMark

        Outputs a check mark and an X.

    .EXAMPLE
        Get-Service | Select-Object Name, CanStop, CanPauseAndContinue |
            ConvertTo-CheckMark -Property CanStop | Format-Table

        Converts only the CanStop property, leaving CanPauseAndContinue as a boolean.

    .EXAMPLE
        $false | ConvertTo-CheckMark -FalseMark ([char]0x2716)

        Uses a plain heavy X (✖) instead of the red X emoji.

    .INPUTS
        System.Boolean, System.Object

    .OUTPUTS
        System.String for boolean input; a copied PSObject for object input.

    .NOTES
        Symbols are specified via [char] code points so the script is safe in
        PowerShell 5.1 regardless of file encoding. If the console shows "??",
        run: [Console]::OutputEncoding = [Text.Encoding]::UTF8
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, Mandatory)]
        $InputObject,

        [string[]]$Property,

        [string]$TrueMark  = [char]0x2705,

        [string]$FalseMark = [char]0x274C
    )
    process {
        if ($null -eq $InputObject) { return $InputObject }

        if ($InputObject -is [bool]) {
            if ($InputObject) { $TrueMark } else { $FalseMark }
            return
        }

        $copy = $InputObject | Select-Object *
        $targets = if ($Property) {
            $copy.PSObject.Properties | Where-Object { $_.Name -in $Property }
        } else {
            $copy.PSObject.Properties
        }

        foreach ($p in $targets) {
            if ($p.Value -is [bool]) {
                $p.Value = if ($p.Value) { $TrueMark } else { $FalseMark }
            }
        }
        $copy
    }
}