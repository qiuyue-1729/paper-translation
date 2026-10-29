param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputDocx,
    [Parameter(Mandatory=$true)][ValidateSet('AxMath','MathType')][string]$Engine,
    [string]$AddinPath,
    [int]$ExpectedCount = 0,
    [string]$OutputPdf,
    [string]$EquationStyle = '公式',
    [ValidateRange(0,72)][double]$FontSize = 0
)

# Run in an interactive Windows user session with Word and a licensed plugin.
# Input must contain only TeX-to-equation conversions, never existing plugin
# equations: TeX Toggle is bidirectional. Work on a new output copy only.
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $InputDocx).Path
$destination = [IO.Path]::GetFullPath($OutputDocx)
if ($source -eq $destination) { throw 'Input and output must be different files.' }
if (Test-Path -LiteralPath $destination) { throw 'Output already exists.' }
$null = New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($destination)) -Force
Copy-Item -LiteralPath $source -Destination $destination
$word = $null
$document = $null
$success = $false
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($destination, $false, $false)
    $before = $document.InlineShapes.Count
    if ($before -gt 0 -or $document.OMaths.Count -gt 0) {
        throw 'Use a staging document without existing shapes or equations; the macro is bidirectional.'
    }
    if ($AddinPath) {
        $addin = (Resolve-Path -LiteralPath $AddinPath).Path
        $null = $word.AddIns.Add($addin, $true)
    }
    $document.Activate()
    $formulaStyle = $document.Styles.Item($EquationStyle)
    if ($Engine -eq 'AxMath') {
        $macro = 'AMCCTex2AMEqn.ConvertTeX2AMEqn'
        $prefix = 'Equation.AxMath'
    } else {
        $macro = 'MTCommand_TeXToggle'
        $prefix = 'Equation.DSMT'
    }
    $pattern = '(?s)\$\$(.+?)\$\$|\\\[(.+?)\\\]|\\\((.+?)\\\)|(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)'
    $segments = [regex]::Matches($document.Content.Text,$pattern)
    if ($segments.Count -eq 0) { throw 'No delimited TeX equations found.' }
    if ($ExpectedCount -gt 0 -and $segments.Count -ne $ExpectedCount) {
        throw "Expected $ExpectedCount delimited equations; found $($segments.Count)."
    }
    for ($i=$segments.Count-1; $i -ge 0; $i--) {
        $segment = $segments[$i]
        $range = $document.Range($segment.Index, $segment.Index+$segment.Length)
        $paragraph = $range.Paragraphs.Item(1).Range
        if ($paragraph.Text.Trim() -ne $segment.Value.Trim()) {
            throw 'Convert standalone equation paragraphs only; keep inline variables in the body style.'
        }
        $paragraph.Style = $EquationStyle
        $paragraph.ParagraphFormat.Reset()
        if ($FontSize -gt 0) { $range.Font.Size = $FontSize }
        $range.Select()
        $oldCount = $document.InlineShapes.Count
        $null = $word.Run($macro)
        if ($document.InlineShapes.Count -ne ($oldCount+1)) {
            throw "Equation conversion failed at source position $($segment.Index)."
        }
    }
    $classes = @()
    foreach ($shape in $document.InlineShapes) {
        $class = $shape.OLEFormat.ProgID
        if ($class -notlike "$prefix*") { throw "Unexpected equation object class: $class" }
        $classes += $class
        # Restore the template equation style after plugin macros add their own formatting.
        $paragraph = $shape.Range.Paragraphs.Item(1).Range
        $paragraph.Style = $EquationStyle
        $paragraph.ParagraphFormat.Reset()
    }
    if ($classes.Count -eq 0) { throw 'Plugin created no equation objects.' }
    if ($ExpectedCount -gt 0 -and $classes.Count -ne $ExpectedCount) {
        throw "Expected $ExpectedCount equations; got $($classes.Count)."
    }
    $document.Save()
    $document.Close(0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    $document = $word.Documents.Open($destination, $false, $true)
    if ($document.InlineShapes.Count -ne $classes.Count) { throw 'Object count changed after reopening.' }
    foreach ($shape in $document.InlineShapes) {
        if ($shape.OLEFormat.ProgID -notlike "$prefix*") { throw 'Plugin class did not survive reopening.' }
    }
    if ($OutputPdf) {
        $pdf = [IO.Path]::GetFullPath($OutputPdf)
        $null = New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($pdf)) -Force
        $document.ExportAsFixedFormat($pdf,17)
    }
    $success = $true
    [pscustomobject]@{ engine=$Engine; macro=$macro; count=$classes.Count; classes=$classes; reopened=$true; output=$destination } | ConvertTo-Json
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    if (-not $success) { Write-Warning 'Conversion failed; output is a diagnostic copy, not an accepted deliverable.' }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
