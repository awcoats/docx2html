<#
.SYNOPSIS
    Converts a Word .docx file into a single self-contained, hierarchical, collapsible HTML file.

.DESCRIPTION
    Expands the .docx (which is a zip archive) into a temp folder and reads word/document.xml,
    word/styles.xml, word/numbering.xml and word/_rels/document.xml.rels directly (no Word/Office
    interop required).

    Headings (Heading 1 - Heading 9, detected via style outline level or style name) are rendered
    as nested <div class="heading-section"> elements, so a Heading 2 that follows a Heading 1
    becomes a DOM child of that Heading 1's section, a Heading 3 becomes a child of that Heading 2,
    and so on. Non-heading paragraphs, list items and tables are placed inside the innermost
    currently-open heading section (or at the document root if no heading has been seen yet).

    List items are detected from direct paragraph numbering (w:pPr/w:numPr) or from a
    paragraph's list style, including Word's built-in "List Bullet"/"List Bullet 2"/"List
    Bullet 3" and "List Number"/"List Number 2"/"List Number 3" styles (whose nesting level is
    inferred from the trailing digit in the style name, since those built-in styles don't
    actually encode w:ilvl themselves) and any custom style that inherits numbering via
    basedOn. They are rendered as properly nested <ul>/<ol> elements based on each paragraph's
    indent level, so a Word multilevel list becomes real nested HTML lists rather than a single
    flat list; if a level is skipped entirely, an empty placeholder <li> is inserted so every
    <ul>/<ol> still only ever directly contains <li> elements (required for valid, reliably
    renderable nested-list HTML). Each level's list is rendered as <ol> or <ul> depending on
    that level's numbering format in word/numbering.xml (bullet/none vs.
    decimal/lowerLetter/upperRoman/etc.).

    A body paragraph is treated as a check box paragraph if it is a Word "Check Box Content
    Control" (w:sdt / w14:checkbox) or its visible text starts with a Unicode check box glyph
    (unchecked U+2610, checked U+2611/U+2612, or the common Wingdings private-use-area
    unchecked/checked glyphs). The glyph/control is replaced with a live, clickable HTML check
    box that always starts unchecked, regardless of the checked state found in the source
    document. Checking it in the browser applies a strike-through style to that paragraph's
    text; unchecking it removes the strike-through. Nothing is ever omitted from the output.

    Two special markers are also recognized in heading text (the marker itself is stripped
    from the displayed heading):
      - A heading containing the literal text "[duplicate]" gets a "Duplicate" button. Clicking
        it clones that heading's entire section (the heading plus everything nested under it)
        and inserts the copy immediately after the original.
      - A heading whose text begins with a "checked" Unicode check box glyph (☑ U+2611, ☒
        U+2612, or the Wingdings private-use-area checked glyph) gets a check box, pre-checked,
        in place of the glyph. Unchecking it hides that heading's entire section (the heading
        plus everything nested under it); checking it again shows it. This is independent of,
        and stacks with, the collapse/expand behavior. Each heading is rendered, in order, as:
        the collapse/expand toggle arrow, then (if present) the visibility check box, then the
        heading text itself.

    The generated HTML embeds plain (no jQuery / no external dependencies) CSS + JavaScript that
    makes every heading clickable to collapse/expand its content, plus "Expand All" / "Collapse All"
    buttons.

.PARAMETER DocxPath
    Path to the source .docx file.

.PARAMETER OutputPath
    Path to write the generated .html file. Defaults to the source file name with a .html extension
    in the same folder.

.PARAMETER Force
    Overwrite OutputPath if it already exists.

.NOTES
    Requires Windows PowerShell 5.1 (uses .NET Framework's System.IO.Compression.FileSystem).
    Does not require Microsoft Word to be installed.

.EXAMPLE
    ./Convert-DocxToHtml.ps1 -DocxPath 'C:\Docs\Design.docx'

.EXAMPLE
    ./Convert-DocxToHtml.ps1 -DocxPath 'C:\Docs\Design.docx' -OutputPath 'C:\Out\design.html' -Force
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [switch]$Force
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

Set-StrictMode -Version Latest

Add-Type -AssemblyName System.IO.Compression.FileSystem

$WNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$RNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
$W14Namespace = "http://schemas.microsoft.com/office/word/2010/wordml"

# Leading glyphs used by hand-typed check box lists: Unicode ballot boxes
# (unchecked / checked / checked-with-x) plus the common Wingdings
# private-use-area unchecked/checked box glyphs.
$UncheckedGlyphs = [char]0x2610, [char]0xF0A8
$CheckedGlyphs = [char]0x2611, [char]0x2612, [char]0xF0FE

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function New-NamespaceManager {
    param([System.Xml.XmlDocument]$XmlDoc)

    $nsMgr = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
    $nsMgr.AddNamespace("w", $WNamespace)
    $nsMgr.AddNamespace("r", $RNamespace)
    $nsMgr.AddNamespace("w14", $W14Namespace)

    # XmlNamespaceManager implements IEnumerable, so PowerShell would otherwise
    # unroll it into its enumerated namespace entries when returned. The comma
    # operator forces it back out as a single object.
    return , $nsMgr
}

function Get-StylesHeadingMap {
    param([string]$StylesXmlPath)

    $map = @{}
    if (-not (Test-Path -LiteralPath $StylesXmlPath)) {
        return $map
    }

    $stylesDoc = New-Object System.Xml.XmlDocument
    $stylesDoc.Load($StylesXmlPath)
    $nsMgr = New-NamespaceManager -XmlDoc $stylesDoc

    $styleNodes = $stylesDoc.SelectNodes("//w:style[@w:type='paragraph']", $nsMgr)
    foreach ($styleNode in $styleNodes) {
        $styleId = $styleNode.GetAttribute("styleId", $WNamespace)
        if ([string]::IsNullOrEmpty($styleId)) { continue }

        $level = $null

        $outlineNode = $styleNode.SelectSingleNode("w:pPr/w:outlineLvl", $nsMgr)
        if ($outlineNode) {
            $val = $outlineNode.GetAttribute("val", $WNamespace)
            if ($val -match '^\d+$') {
                $level = [int]$val + 1
            }
        }

        if (-not $level) {
            $nameNode = $styleNode.SelectSingleNode("w:name", $nsMgr)
            if ($nameNode) {
                $name = $nameNode.GetAttribute("val", $WNamespace)
                if ($name -match '(?i)^heading\s*([1-9])$') {
                    $level = [int]$Matches[1]
                }
            }
        }

        if ($level -and $level -ge 1 -and $level -le 9) {
            $map[$styleId] = $level
        }
    }

    return $map
}

function Get-ParagraphNumPrInfo {
    param(
        [System.Xml.XmlElement]$NumPrNode,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $numIdNode = $NumPrNode.SelectSingleNode("w:numId", $NsMgr)
    $numId = $null
    if ($numIdNode) {
        $numId = $numIdNode.GetAttribute("val", $WNamespace)
    }
    if ([string]::IsNullOrEmpty($numId)) {
        return $null
    }

    $ilvl = 0
    $ilvlNode = $NumPrNode.SelectSingleNode("w:ilvl", $NsMgr)
    if ($ilvlNode) {
        $ilvlVal = $ilvlNode.GetAttribute("val", $WNamespace)
        if ($ilvlVal -match '^\d+$') {
            $ilvl = [int]$ilvlVal
        }
    }

    return @{ NumId = $numId; Ilvl = $ilvl }
}

function Get-StylesListInfoMap {
    param([string]$StylesXmlPath)

    $result = @{}
    if (-not (Test-Path -LiteralPath $StylesXmlPath)) {
        return $result
    }

    $stylesDoc = New-Object System.Xml.XmlDocument
    $stylesDoc.Load($StylesXmlPath)
    $nsMgr = New-NamespaceManager -XmlDoc $stylesDoc

    # A style is a "list style" if it (or a style it's basedOn, walking the
    # chain) directly defines a numPr. Word's built-in "List Bullet"/"List
    # Number" styles carry numPr (and an indent/level) on the style itself,
    # not on each paragraph.
    $ownNumPr = @{}
    $basedOn = @{}
    $allStyleIds = New-Object System.Collections.Generic.List[string]

    $styleNodes = $stylesDoc.SelectNodes("//w:style[@w:type='paragraph']", $nsMgr)
    foreach ($styleNode in $styleNodes) {
        $styleId = $styleNode.GetAttribute("styleId", $WNamespace)
        if ([string]::IsNullOrEmpty($styleId)) { continue }
        [void]$allStyleIds.Add($styleId)

        $numPrNode = $styleNode.SelectSingleNode("w:pPr/w:numPr", $nsMgr)
        if ($numPrNode) {
            $numPrInfo = Get-ParagraphNumPrInfo -NumPrNode $numPrNode -NsMgr $nsMgr
            if ($numPrInfo) {
                # Word's built-in "List Bullet"/"List Bullet 2"/"List Bullet 3" (and the
                # "List Number" equivalents) each carry their OWN numId in styles.xml, but
                # none of them actually set w:ilvl there - it's always absent/0. The trailing
                # digit in the style name is the only signal for how deep that level is, so
                # it takes priority over whatever ilvl (always 0) the style's numPr reported.
                $nameNode = $styleNode.SelectSingleNode("w:name", $nsMgr)
                if ($nameNode) {
                    $name = $nameNode.GetAttribute("val", $WNamespace)
                    if ($name -match '(?i)^list\s+(bullet|number)\s*(\d*)$') {
                        if ([string]::IsNullOrEmpty($Matches[2])) {
                            $numPrInfo.Ilvl = 0
                        }
                        else {
                            $numPrInfo.Ilvl = [int]$Matches[2] - 1
                        }
                    }
                }
                $ownNumPr[$styleId] = $numPrInfo
            }
        }

        $basedOnNode = $styleNode.SelectSingleNode("w:basedOn", $nsMgr)
        if ($basedOnNode) {
            $basedOn[$styleId] = $basedOnNode.GetAttribute("val", $WNamespace)
        }
    }

    foreach ($styleId in $allStyleIds) {
        $current = $styleId
        $visited = New-Object 'System.Collections.Generic.HashSet[string]'
        while ($current -and $visited.Add($current)) {
            if ($ownNumPr.ContainsKey($current)) {
                $result[$styleId] = $ownNumPr[$current]
                break
            }
            $current = $basedOn[$current]
        }
    }

    return $result
}

function Get-NumberingFormatMap {
    param([string]$NumberingXmlPath)

    $result = @{
        NumIdToAbstractId   = @{}
        AbstractLevelFormats = @{}
    }

    if (-not (Test-Path -LiteralPath $NumberingXmlPath)) {
        return $result
    }

    $numDoc = New-Object System.Xml.XmlDocument
    $numDoc.Load($NumberingXmlPath)
    $nsMgr = New-NamespaceManager -XmlDoc $numDoc

    $numNodes = $numDoc.SelectNodes("//w:num", $nsMgr)
    foreach ($numNode in $numNodes) {
        $numId = $numNode.GetAttribute("numId", $WNamespace)
        if ([string]::IsNullOrEmpty($numId)) { continue }

        $abstractNode = $numNode.SelectSingleNode("w:abstractNumId", $nsMgr)
        if ($abstractNode) {
            $abstractId = $abstractNode.GetAttribute("val", $WNamespace)
            if (-not [string]::IsNullOrEmpty($abstractId)) {
                $result.NumIdToAbstractId[$numId] = $abstractId
            }
        }
    }

    $abstractNodes = $numDoc.SelectNodes("//w:abstractNum", $nsMgr)
    foreach ($abstractNode in $abstractNodes) {
        $abstractId = $abstractNode.GetAttribute("abstractNumId", $WNamespace)
        if ([string]::IsNullOrEmpty($abstractId)) { continue }

        $levelMap = @{}
        $lvlNodes = $abstractNode.SelectNodes("w:lvl", $nsMgr)
        foreach ($lvlNode in $lvlNodes) {
            $ilvlAttr = $lvlNode.GetAttribute("ilvl", $WNamespace)
            if ($ilvlAttr -notmatch '^\d+$') { continue }
            $ilvl = [int]$ilvlAttr

            $numFmtNode = $lvlNode.SelectSingleNode("w:numFmt", $nsMgr)
            if ($numFmtNode) {
                $fmt = $numFmtNode.GetAttribute("val", $WNamespace)
                if (-not [string]::IsNullOrEmpty($fmt)) {
                    $levelMap[$ilvl] = $fmt
                }
            }
        }
        $result.AbstractLevelFormats[$abstractId] = $levelMap
    }

    return $result
}

function Get-ParagraphListInfo {
    param(
        [System.Xml.XmlElement]$ParagraphNode,
        [System.Xml.XmlNamespaceManager]$NsMgr,
        [hashtable]$StylesListInfoMap,
        [hashtable]$NumberingFormatMap
    )

    $numPrInfo = $null

    $numPrNode = $ParagraphNode.SelectSingleNode("w:pPr/w:numPr", $NsMgr)
    if ($numPrNode) {
        $numPrInfo = Get-ParagraphNumPrInfo -NumPrNode $numPrNode -NsMgr $NsMgr
    }

    if (-not $numPrInfo) {
        $pStyleNode = $ParagraphNode.SelectSingleNode("w:pPr/w:pStyle", $NsMgr)
        if ($pStyleNode) {
            $styleId = $pStyleNode.GetAttribute("val", $WNamespace)
            if (-not [string]::IsNullOrEmpty($styleId) -and $StylesListInfoMap.ContainsKey($styleId)) {
                $numPrInfo = $StylesListInfoMap[$styleId]
            }
        }
    }

    if (-not $numPrInfo) {
        return $null
    }

    $numFmt = $null
    if ($NumberingFormatMap.NumIdToAbstractId.ContainsKey($numPrInfo.NumId)) {
        $abstractId = $NumberingFormatMap.NumIdToAbstractId[$numPrInfo.NumId]
        if ($NumberingFormatMap.AbstractLevelFormats.ContainsKey($abstractId)) {
            $levelMap = $NumberingFormatMap.AbstractLevelFormats[$abstractId]
            if ($levelMap.ContainsKey($numPrInfo.Ilvl)) {
                $numFmt = $levelMap[$numPrInfo.Ilvl]
            }
        }
    }

    $isOrdered = ($numFmt -and $numFmt -ne "bullet" -and $numFmt -ne "none")

    return @{ Level = $numPrInfo.Ilvl; Ordered = $isOrdered }
}

function Get-RelationshipMap {
    param([string]$RelsXmlPath)

    $map = @{}
    if (-not (Test-Path -LiteralPath $RelsXmlPath)) {
        return $map
    }

    $relsDoc = New-Object System.Xml.XmlDocument
    $relsDoc.Load($RelsXmlPath)

    foreach ($rel in $relsDoc.DocumentElement.ChildNodes) {
        if ($rel.LocalName -eq 'Relationship') {
            $id = $rel.GetAttribute("Id")
            $target = $rel.GetAttribute("Target")
            if (-not [string]::IsNullOrEmpty($id)) {
                $map[$id] = $target
            }
        }
    }

    return $map
}

function Get-HeadingLevel {
    param(
        [System.Xml.XmlElement]$ParagraphNode,
        [System.Xml.XmlNamespaceManager]$NsMgr,
        [hashtable]$StylesMap
    )

    $pPr = $ParagraphNode.SelectSingleNode("w:pPr", $NsMgr)
    if (-not $pPr) { return $null }

    # A direct outline-level override on the paragraph itself wins.
    $outlineNode = $pPr.SelectSingleNode("w:outlineLvl", $NsMgr)
    if ($outlineNode) {
        $val = $outlineNode.GetAttribute("val", $WNamespace)
        if ($val -match '^\d+$') {
            return [int]$val + 1
        }
    }

    $styleNode = $pPr.SelectSingleNode("w:pStyle", $NsMgr)
    if ($styleNode) {
        $styleId = $styleNode.GetAttribute("val", $WNamespace)
        if (-not [string]::IsNullOrEmpty($styleId) -and $StylesMap.ContainsKey($styleId)) {
            return $StylesMap[$styleId]
        }
    }

    return $null
}

function Get-ParagraphRawText {
    param(
        [System.Xml.XmlElement]$ParagraphNode,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $sb = New-Object System.Text.StringBuilder
    $textNodes = $ParagraphNode.SelectNodes(".//w:t", $NsMgr)
    foreach ($textNode in $textNodes) {
        [void]$sb.Append($textNode.InnerText)
    }

    return $sb.ToString()
}

function Get-ParagraphCheckboxInfo {
    param(
        [System.Xml.XmlElement]$ParagraphNode,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    # Word "Check Box Content Control" (Developer tab): a w:sdt whose
    # w:sdtPr carries a w14:checkbox with an explicit checked state.
    $sdtNode = $ParagraphNode.SelectSingleNode(".//w:sdt[w:sdtPr/w14:checkbox]", $NsMgr)
    if ($sdtNode) {
        $isChecked = $false
        $checkedNode = $sdtNode.SelectSingleNode("w:sdtPr/w14:checkbox/w14:checked", $NsMgr)
        if ($checkedNode) {
            $val = $checkedNode.GetAttribute("val", $W14Namespace)
            $isChecked = ($val -eq "1" -or $val -eq "true")
        }

        $info = @{
            IsCheckbox = $true
            Checked    = $isChecked
            SkipNode   = $sdtNode
        }
        return $info
    }

    # Hand-typed check box list: the paragraph's visible text starts with a
    # recognized check box glyph.
    $rawText = (Get-ParagraphRawText -ParagraphNode $ParagraphNode -NsMgr $NsMgr).TrimStart()
    if ($rawText.Length -gt 0) {
        $firstChar = $rawText[0]

        if ($UncheckedGlyphs -contains $firstChar) {
            $info = @{
                IsCheckbox = $true
                Checked    = $false
                SkipNode   = $null
            }
            return $info
        }

        if ($CheckedGlyphs -contains $firstChar) {
            $info = @{
                IsCheckbox = $true
                Checked    = $true
                SkipNode   = $null
            }
            return $info
        }
    }

    return $null
}

function Test-ToggleProperty {
    param(
        [System.Xml.XmlElement]$RunPropertiesNode,
        [string]$TagName,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    if (-not $RunPropertiesNode) { return $false }

    $node = $RunPropertiesNode.SelectSingleNode("w:$TagName", $NsMgr)
    if (-not $node) { return $false }

    $val = $node.GetAttribute("val", $WNamespace)
    if ([string]::IsNullOrEmpty($val)) { return $true }

    return ($val -ne "0" -and $val -ne "false")
}

function Get-RunHtml {
    param(
        [System.Xml.XmlElement]$RunNode,
        [System.Xml.XmlNamespaceManager]$NsMgr
    )

    $textBuilder = New-Object System.Text.StringBuilder
    foreach ($child in $RunNode.ChildNodes) {
        switch ($child.LocalName) {
            't' { [void]$textBuilder.Append($child.InnerText) }
            'tab' { [void]$textBuilder.Append("`t") }
            'br' { [void]$textBuilder.Append("`n") }
            'cr' { [void]$textBuilder.Append("`n") }
            default { }
        }
    }

    $rawText = $textBuilder.ToString()
    if ([string]::IsNullOrEmpty($rawText)) { return "" }

    $encoded = [System.Net.WebUtility]::HtmlEncode($rawText) -replace "`n", "<br/>"

    $rPr = $RunNode.SelectSingleNode("w:rPr", $NsMgr)
    if (Test-ToggleProperty -RunPropertiesNode $rPr -TagName "b" -NsMgr $NsMgr) {
        $encoded = "<strong>$encoded</strong>"
    }
    if (Test-ToggleProperty -RunPropertiesNode $rPr -TagName "i" -NsMgr $NsMgr) {
        $encoded = "<em>$encoded</em>"
    }
    if (Test-ToggleProperty -RunPropertiesNode $rPr -TagName "u" -NsMgr $NsMgr) {
        $encoded = "<u>$encoded</u>"
    }

    return $encoded
}

function Get-InlineHtml {
    param(
        [System.Xml.XmlElement]$ContainerNode,
        [System.Xml.XmlNamespaceManager]$NsMgr,
        [hashtable]$RelsMap,
        [System.Xml.XmlElement]$SkipNode = $null
    )

    $sb = New-Object System.Text.StringBuilder
    foreach ($child in $ContainerNode.ChildNodes) {
        switch ($child.LocalName) {
            'r' {
                [void]$sb.Append((Get-RunHtml -RunNode $child -NsMgr $NsMgr))
            }
            'hyperlink' {
                $rId = $child.GetAttribute("id", $RNamespace)
                $href = "#"
                if (-not [string]::IsNullOrEmpty($rId) -and $RelsMap.ContainsKey($rId)) {
                    $href = $RelsMap[$rId]
                }
                $inner = Get-InlineHtml -ContainerNode $child -NsMgr $NsMgr -RelsMap $RelsMap -SkipNode $SkipNode
                if (-not [string]::IsNullOrEmpty($inner)) {
                    $hrefEncoded = [System.Net.WebUtility]::HtmlEncode($href)
                    [void]$sb.Append("<a href=`"$hrefEncoded`" target=`"_blank`" rel=`"noopener`">$inner</a>")
                }
            }
            'smartTag' {
                [void]$sb.Append((Get-InlineHtml -ContainerNode $child -NsMgr $NsMgr -RelsMap $RelsMap -SkipNode $SkipNode))
            }
            'sdt' {
                # A recognized check box content control is rendered separately
                # as an HTML <input>, so its own glyph text is left out here.
                $isSkipped = $false
                if ($SkipNode) {
                    $isSkipped = [object]::ReferenceEquals($child, $SkipNode)
                }
                if (-not $isSkipped) {
                    $sdtContentNode = $child.SelectSingleNode("w:sdtContent", $NsMgr)
                    if ($sdtContentNode) {
                        [void]$sb.Append((Get-InlineHtml -ContainerNode $sdtContentNode -NsMgr $NsMgr -RelsMap $RelsMap -SkipNode $SkipNode))
                    }
                }
            }
            default { }
        }
    }

    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Nested list rendering.
#
# $ListStack is a System.Collections.Generic.List of @{ Level = <int>;
# Ordered = <bool> } frames, shallowest first / deepest last. The invariant
# maintained between calls is: whenever the stack is non-empty, the deepest
# (top) frame always has one currently-open, not-yet-closed <li> waiting for
# either more inline content, a nested list, or a closing </li>. That is what
# lets a deeper list be opened *inside* the parent's still-open <li>, which is
# what produces valid nested <ul>/<ol> markup.
# ---------------------------------------------------------------------------

function Close-ListLevel {
    param(
        [System.Text.StringBuilder]$Content,
        [System.Collections.Generic.List[object]]$ListStack
    )

    $top = $ListStack[$ListStack.Count - 1]
    $ListStack.RemoveAt($ListStack.Count - 1)

    [void]$Content.Append("</li>`n")
    if ($top.Ordered) {
        [void]$Content.Append("</ol>`n")
    }
    else {
        [void]$Content.Append("</ul>`n")
    }
}

function Close-AllListLevels {
    param(
        [System.Text.StringBuilder]$Content,
        [System.Collections.Generic.List[object]]$ListStack
    )

    while ($ListStack.Count -gt 0) {
        Close-ListLevel -Content $Content -ListStack $ListStack
    }
}

function Add-ListItem {
    param(
        [System.Text.StringBuilder]$Content,
        [System.Collections.Generic.List[object]]$ListStack,
        [int]$Level,
        [bool]$Ordered,
        [string]$ItemHtml
    )

    # Close any open, deeper levels first so the stack's remaining top frame
    # (if any) has a level <= the item we're about to add.
    while ($ListStack.Count -gt 0 -and $ListStack[$ListStack.Count - 1].Level -gt $Level) {
        Close-ListLevel -Content $Content -ListStack $ListStack
    }

    $startLevel = -1
    if ($ListStack.Count -gt 0) {
        $startLevel = $ListStack[$ListStack.Count - 1].Level
    }

    if ($startLevel -lt $Level) {
        # Open one nested <ul>/<ol> per missing level, deepest last. When the
        # stack wasn't empty, this happens inside the parent's still-open
        # <li>, which is exactly the markup nested lists need. If more than
        # one level is being skipped at once (e.g. going straight from an
        # empty stack to Level 2), the intermediate levels get an empty
        # placeholder <li> so every <ul>/<ol> we open still only ever
        # directly contains <li> elements, which is required for valid HTML.
        for ($lvl = $startLevel + 1; $lvl -le $Level; $lvl++) {
            if ($Ordered) {
                [void]$Content.Append("`n<ol>`n")
            }
            else {
                [void]$Content.Append("`n<ul>`n")
            }
            $frame = [PSCustomObject]@{ Level = $lvl; Ordered = $Ordered }
            [void]$ListStack.Add($frame)

            if ($lvl -lt $Level) {
                [void]$Content.Append("  <li>")
            }
        }
    }
    else {
        # Same level as the currently open list: close its pending <li>. If
        # the list type (bullet vs. numbered) changed at this same level,
        # swap the wrapper element too.
        $top = $ListStack[$ListStack.Count - 1]
        if ($top.Ordered -ne $Ordered) {
            $ListStack.RemoveAt($ListStack.Count - 1)
            [void]$Content.Append("</li>`n")
            if ($top.Ordered) {
                [void]$Content.Append("</ol>`n")
            }
            else {
                [void]$Content.Append("</ul>`n")
            }
            if ($Ordered) {
                [void]$Content.Append("<ol>`n")
            }
            else {
                [void]$Content.Append("<ul>`n")
            }
            $frame = [PSCustomObject]@{ Level = $Level; Ordered = $Ordered }
            [void]$ListStack.Add($frame)
        }
        else {
            [void]$Content.Append("</li>`n")
        }
    }

    [void]$Content.Append("  <li>$ItemHtml")
}

function Get-TableHtml {
    param(
        [System.Xml.XmlElement]$TableNode,
        [System.Xml.XmlNamespaceManager]$NsMgr,
        [hashtable]$RelsMap
    )

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table class=`"docx-table`">`n")

    $rowNodes = $TableNode.SelectNodes("w:tr", $NsMgr)
    foreach ($rowNode in $rowNodes) {
        [void]$sb.Append("  <tr>`n")
        $cellNodes = $rowNode.SelectNodes("w:tc", $NsMgr)
        foreach ($cellNode in $cellNodes) {
            $cellParts = New-Object System.Collections.Generic.List[string]
            $paragraphNodes = $cellNode.SelectNodes("w:p", $NsMgr)
            foreach ($paragraphNode in $paragraphNodes) {
                $text = Get-InlineHtml -ContainerNode $paragraphNode -NsMgr $NsMgr -RelsMap $RelsMap
                if (-not [string]::IsNullOrEmpty($text)) {
                    $cellParts.Add($text)
                }
            }
            $cellHtml = [string]::Join("<br/>", $cellParts)
            [void]$sb.Append("    <td>$cellHtml</td>`n")
        }
        [void]$sb.Append("  </tr>`n")
    }

    [void]$sb.Append("</table>`n")
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$resolvedDocx = Resolve-Path -LiteralPath $DocxPath -ErrorAction Stop
$DocxPath = $resolvedDocx.ProviderPath

if (-not $OutputPath) {
    $OutputPath = [System.IO.Path]::ChangeExtension($DocxPath, ".html")
}

if ((Test-Path -LiteralPath $OutputPath) -and (-not $Force)) {
    throw "Output file '$OutputPath' already exists. Use -Force to overwrite."
}

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("docx2html_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    Write-Information "Expanding '$DocxPath' ..."
    [System.IO.Compression.ZipFile]::ExtractToDirectory($DocxPath, $tempDir)

    $documentXmlPath = Join-Path $tempDir "word\document.xml"
    $stylesXmlPath = Join-Path $tempDir "word\styles.xml"
    $relsXmlPath = Join-Path $tempDir "word\_rels\document.xml.rels"
    $numberingXmlPath = Join-Path $tempDir "word\numbering.xml"

    if (-not (Test-Path -LiteralPath $documentXmlPath)) {
        throw "'$DocxPath' does not look like a valid .docx file (word/document.xml not found)."
    }

    $stylesMap = Get-StylesHeadingMap -StylesXmlPath $stylesXmlPath
    $stylesListInfoMap = Get-StylesListInfoMap -StylesXmlPath $stylesXmlPath
    $numberingFormatMap = Get-NumberingFormatMap -NumberingXmlPath $numberingXmlPath
    $relsMap = Get-RelationshipMap -RelsXmlPath $relsXmlPath

    $documentDoc = New-Object System.Xml.XmlDocument
    $documentDoc.Load($documentXmlPath)
    $nsMgr = New-NamespaceManager -XmlDoc $documentDoc

    $bodyNode = $documentDoc.SelectSingleNode("//w:body", $nsMgr)
    if (-not $bodyNode) {
        throw "Could not find <w:body> in word/document.xml."
    }

    $content = New-Object System.Text.StringBuilder
    $openLevels = New-Object System.Collections.Generic.List[int]
    $listStack = New-Object System.Collections.Generic.List[object]

    foreach ($node in $bodyNode.ChildNodes) {
        switch ($node.LocalName) {

            'p' {
                $level = Get-HeadingLevel -ParagraphNode $node -NsMgr $nsMgr -StylesMap $stylesMap

                if ($level) {
                    Close-AllListLevels -Content $content -ListStack $listStack

                    while ($openLevels.Count -gt 0 -and $openLevels[$openLevels.Count - 1] -ge $level) {
                        [void]$content.Append("</div></div>`n")
                        $openLevels.RemoveAt($openLevels.Count - 1)
                    }

                    $rawHeadingText = Get-ParagraphRawText -ParagraphNode $node -NsMgr $nsMgr
                    $isDuplicateHeading = $rawHeadingText -match '(?i)\[duplicate\]'

                    # The check box glyph doesn't have to be the very first character - Google
                    # Docs/Word exports often prepend manual outline numbering text (e.g. "1 ",
                    # "1.2 ") to the heading before the glyph - so search the whole heading
                    # rather than only checking the first character.
                    $isVisibilityHeading = $rawHeadingText -match "[\u2611\u2612\uF0FE]"

                    $headingHtml = Get-InlineHtml -ContainerNode $node -NsMgr $nsMgr -RelsMap $relsMap
                    if ($isDuplicateHeading) {
                        $headingHtml = ($headingHtml -replace '(?i)\s*\[duplicate\]\s*', ' ').Trim()
                    }
                    if ($isVisibilityHeading) {
                        $headingHtml = ($headingHtml -replace "[\u2611\u2612\uF0FE]\s*", '').Trim()
                    }
                    if ([string]::IsNullOrWhiteSpace($headingHtml)) {
                        $headingHtml = "&nbsp;"
                    }
                    $hTag = "h" + [Math]::Min($level, 6)

                    [void]$content.Append("<div class=`"heading-section`" data-level=`"$level`">`n")
                    [void]$content.Append("  <div class=`"heading-header`" onclick=`"toggleSection(this)`">`n")
                    [void]$content.Append("    <span class=`"toggle-icon`">&#9662;</span>`n")
                    if ($isVisibilityHeading) {
                        [void]$content.Append("    <input type=`"checkbox`" class=`"visibility-toggle`" checked title=`"Show/hide this section`" onclick=`"event.stopPropagation()`" onchange=`"toggleVisibilitySection(this)`"/>`n")
                    }
                    [void]$content.Append("    <$hTag>$headingHtml</$hTag>`n")
                    if ($isDuplicateHeading) {
                        [void]$content.Append("    <button type=`"button`" class=`"duplicate-btn`" onclick=`"event.stopPropagation(); duplicateSection(this)`">Duplicate</button>`n")
                    }
                    [void]$content.Append("  </div>`n")
                    [void]$content.Append("  <div class=`"heading-content`">`n")

                    $openLevels.Add($level)
                }
                else {
                    # A paragraph can be BOTH a check box (glyph/content-control prefix) AND a
                    # list item (numPr / list style) at the same time - e.g. a bulleted to-do
                    # list where every item starts with a check box glyph. Check-box detection
                    # only decides how the paragraph's own inline content is rendered (plain
                    # text vs. an interactive check box); list-item detection independently
                    # decides the wrapping element (<li> inside a nested list vs. a standalone
                    # <p>), so neither one silently discards the other.
                    $checkboxInfo = Get-ParagraphCheckboxInfo -ParagraphNode $node -NsMgr $nsMgr
                    $listInfo = Get-ParagraphListInfo -ParagraphNode $node -NsMgr $nsMgr -StylesListInfoMap $stylesListInfoMap -NumberingFormatMap $numberingFormatMap

                    if ($checkboxInfo) {
                        $labelHtml = Get-InlineHtml -ContainerNode $node -NsMgr $nsMgr -RelsMap $relsMap -SkipNode $checkboxInfo.SkipNode
                        $labelHtml = $labelHtml -replace "^[\u2610\u2611\u2612\uF0A8\uF0FE]\s*", ""
                        if ([string]::IsNullOrWhiteSpace($labelHtml)) {
                            $labelHtml = "&nbsp;"
                        }

                        # Paragraph check boxes always render unchecked with no
                        # strike-through, regardless of the glyph found in the source
                        # document; strike-through only ever reflects a reader's own click.
                        $itemHtml = "<label class=`"checkbox-item-label`"><input type=`"checkbox`" class=`"strike-checkbox`" onchange=`"toggleStrike(this)`"/> <span class=`"strike-text`">$labelHtml</span></label>"
                    }
                    else {
                        $itemHtml = Get-InlineHtml -ContainerNode $node -NsMgr $nsMgr -RelsMap $relsMap
                    }

                    if ($listInfo) {
                        if ([string]::IsNullOrEmpty($itemHtml)) {
                            $itemHtml = "&nbsp;"
                        }
                        Add-ListItem -Content $content -ListStack $listStack -Level $listInfo.Level -Ordered $listInfo.Ordered -ItemHtml $itemHtml
                    }
                    elseif ($checkboxInfo) {
                        Close-AllListLevels -Content $content -ListStack $listStack
                        [void]$content.Append("<p class=`"checkbox-paragraph`">$itemHtml</p>`n")
                    }
                    else {
                        Close-AllListLevels -Content $content -ListStack $listStack
                        if (-not [string]::IsNullOrEmpty($itemHtml)) {
                            [void]$content.Append("<p>$itemHtml</p>`n")
                        }
                    }
                }
            }

            'tbl' {
                Close-AllListLevels -Content $content -ListStack $listStack
                [void]$content.Append((Get-TableHtml -TableNode $node -NsMgr $nsMgr -RelsMap $relsMap))
            }

            default {
                # sectPr and anything else (e.g. bookmarks) are intentionally ignored.
            }
        }
    }

    Close-AllListLevels -Content $content -ListStack $listStack
    while ($openLevels.Count -gt 0) {
        [void]$content.Append("</div></div>`n")
        $openLevels.RemoveAt($openLevels.Count - 1)
    }

    $docTitle = [System.Net.WebUtility]::HtmlEncode([System.IO.Path]::GetFileNameWithoutExtension($DocxPath))

    $css = @'
<style>
  :root {
    --border-color: #d9d9d9;
    --header-bg: #f5f6f8;
    --header-hover-bg: #eaecef;
    --accent: #2b579a;
  }
  body {
    font-family: "Segoe UI", Calibri, Arial, sans-serif;
    color: #202020;
    max-width: 960px;
    margin: 2rem auto;
    padding: 0 1rem 4rem 1rem;
    line-height: 1.5;
  }
  .toolbar {
    position: sticky;
    top: 0;
    background: #fff;
    padding: 0.5rem 0;
    margin-bottom: 1rem;
    border-bottom: 1px solid var(--border-color);
    z-index: 10;
  }
  .toolbar button {
    font-size: 0.85rem;
    padding: 0.35rem 0.8rem;
    margin-right: 0.5rem;
    border: 1px solid var(--border-color);
    border-radius: 4px;
    background: var(--header-bg);
    cursor: pointer;
  }
  .toolbar button:hover {
    background: var(--header-hover-bg);
  }
  .heading-section {
    border-left: 2px solid var(--border-color);
    margin: 0.25rem 0 0.25rem 0.1rem;
  }
  .heading-header {
    display: flex;
    align-items: center;
    gap: 0.4rem;
    cursor: pointer;
    background: var(--header-bg);
    padding: 0.25rem 0.5rem;
    border-radius: 4px;
    user-select: none;
  }
  .heading-header:hover {
    background: var(--header-hover-bg);
  }
  .heading-header h1, .heading-header h2, .heading-header h3,
  .heading-header h4, .heading-header h5, .heading-header h6 {
    margin: 0.2rem 0;
    color: var(--accent);
  }
  .toggle-icon {
    display: inline-block;
    transition: transform 0.15s ease-in-out;
    color: var(--accent);
    flex: 0 0 auto;
  }
  .heading-content {
    padding-left: 1.25rem;
    border-left: 1px dashed var(--border-color);
    margin-left: 0.4rem;
  }
  .heading-section.collapsed > .heading-content {
    display: none;
  }
  .heading-section.collapsed > .heading-header .toggle-icon {
    transform: rotate(-90deg);
  }
  .heading-section.hidden-section > .heading-content {
    display: none;
  }
  .heading-header .visibility-toggle {
    flex: 0 0 auto;
    cursor: pointer;
  }
  .heading-header .duplicate-btn {
    flex: 0 0 auto;
    margin-left: auto;
    font-size: 0.75rem;
    padding: 0.2rem 0.6rem;
    border: 1px solid var(--accent);
    border-radius: 4px;
    background: #fff;
    color: var(--accent);
    cursor: pointer;
  }
  .heading-header .duplicate-btn:hover {
    background: var(--accent);
    color: #fff;
  }
  table.docx-table {
    border-collapse: collapse;
    margin: 0.75rem 0;
    width: 100%;
  }
  table.docx-table td {
    border: 1px solid var(--border-color);
    padding: 0.4rem 0.6rem;
    vertical-align: top;
  }
  ul, ol {
    margin: 0.4rem 0;
  }
  li > ul, li > ol {
    margin: 0.1rem 0 0.1rem 0.5rem;
  }
  p.checkbox-paragraph {
    margin: 0.3rem 0;
  }
  .checkbox-item-label {
    display: flex;
    align-items: flex-start;
    gap: 0.4rem;
    cursor: pointer;
  }
  .checkbox-item-label input.strike-checkbox {
    margin-top: 0.2rem;
    flex: 0 0 auto;
    cursor: pointer;
  }
  li > .checkbox-item-label {
    display: inline-flex;
  }
  .strike-text.struck {
    text-decoration: line-through;
    opacity: 0.6;
  }
</style>
'@

    $js = @'
<script>
  function toggleSection(headerEl) {
    var section = headerEl.parentElement;
    section.classList.toggle('collapsed');
  }

  function setAllSections(collapsed) {
    var sections = document.querySelectorAll('.heading-section');
    for (var i = 0; i < sections.length; i++) {
      if (collapsed) {
        sections[i].classList.add('collapsed');
      } else {
        sections[i].classList.remove('collapsed');
      }
    }
  }

  function duplicateSection(buttonEl) {
    var section = buttonEl.closest('.heading-section');
    if (!section) {
      return;
    }
    var clone = section.cloneNode(true);
    section.parentNode.insertBefore(clone, section.nextSibling);
  }

  function toggleVisibilitySection(checkboxEl) {
    var section = checkboxEl.closest('.heading-section');
    if (!section) {
      return;
    }
    if (checkboxEl.checked) {
      section.classList.remove('hidden-section');
    } else {
      section.classList.add('hidden-section');
    }
  }

  function toggleStrike(checkboxEl) {
    var label = checkboxEl.closest('label');
    var textEl = label ? label.querySelector('.strike-text') : null;
    if (!textEl) {
      return;
    }
    if (checkboxEl.checked) {
      textEl.classList.add('struck');
    } else {
      textEl.classList.remove('struck');
    }
  }
</script>
'@

    $html = New-Object System.Text.StringBuilder
    [void]$html.Append("<!DOCTYPE html>`n")
    [void]$html.Append("<html lang=`"en`">`n<head>`n")
    [void]$html.Append("<meta charset=`"utf-8`"/>`n")
    [void]$html.Append("<meta name=`"viewport`" content=`"width=device-width, initial-scale=1`"/>`n")
    [void]$html.Append("<title>$docTitle</title>`n")
    [void]$html.Append($css)
    [void]$html.Append("`n</head>`n<body>`n")
    [void]$html.Append("<div class=`"toolbar`">`n")
    [void]$html.Append("  <button type=`"button`" onclick=`"setAllSections(false)`">Expand All</button>`n")
    [void]$html.Append("  <button type=`"button`" onclick=`"setAllSections(true)`">Collapse All</button>`n")
    [void]$html.Append("</div>`n")
    [void]$html.Append("<div class=`"document-root`">`n")
    [void]$html.Append($content.ToString())
    [void]$html.Append("</div>`n")
    [void]$html.Append($js)
    [void]$html.Append("`n</body>`n</html>`n")

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($OutputPath, $html.ToString(), $utf8NoBom)

    Write-Information "Wrote '$OutputPath'."
    Get-Item -LiteralPath $OutputPath
}
finally {
    if (Test-Path -LiteralPath $tempDir) {
        Remove-Item -LiteralPath $tempDir -Recurse -Force
    }
}
