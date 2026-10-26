[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [string]$TargetPath,
    [ValidateSet('Auto','All','Omml','Latex','Images')][string]$Mode = 'Auto',
    [ValidateSet('Delimited','Heuristic')][string]$LatexMode = 'Delimited',
    [string]$SectionHeading,
    [string]$EndHeading,
    [ValidateSet('Auto','AltText','All','None')][string]$ImagePolicy = 'Auto',
    [string]$ImageNumbers,
    [ValidateSet('Vlm','Mathpix','Manual')][string]$OcrBackend = 'Vlm',
    [string]$OcrModel,
    [string]$OcrBaseUrl,
    [string]$OcrApiKey,
    [string]$ManualOcrMapPath,
    [double]$MinOcrConfidence = 0.65,
    [int]$MaxFormulas = 0,
    [switch]$DryRun,
    [switch]$KeepWorkFiles,
    [string]$WorkDir,
    [string]$PandocPath,
    [string]$PythonExe = 'python',
    [switch]$Force,
    [switch]$SkipPreflight
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$extractScript = Join-Path $scriptDir 'extract_math_map.py'
$ocrScript = Join-Path $scriptDir 'ocr_formula_image.py'
$setupScript = Join-Path $scriptDir 'setup_environment.ps1'
$preflightRuns = 0
$wordFormatDocx = 16
$wdAlignParagraphCenter = 1
$sourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
$sourceExt = [IO.Path]::GetExtension($sourcePath).ToLowerInvariant()

if ([string]::IsNullOrWhiteSpace($TargetPath)) {
    $base = [IO.Path]::GetFileNameWithoutExtension($sourcePath)
    $TargetPath = Join-Path (Split-Path -Parent $sourcePath) ($base + '_mathtype.docx')
}
$targetPath = [IO.Path]::GetFullPath($TargetPath)
if (-not $DryRun -and $sourceExt -eq '.docx' -and $sourcePath -eq $targetPath) {
    throw 'TargetPath must differ from SourcePath. Work on an output copy.'
}
if (-not $DryRun -and (Test-Path -LiteralPath $targetPath) -and -not $Force) {
    throw "Target exists. Pass -Force to overwrite: $targetPath"
}

$initialRequireOcr = @('Images','All') -contains $Mode
if ($initialRequireOcr -and $OcrBackend -eq 'Manual' -and [string]::IsNullOrWhiteSpace($ManualOcrMapPath)) {
    throw 'Manual OCR requires -ManualOcrMapPath.'
}

function Assert-FileClosed {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $stream.Close()
    } catch {
        throw "$Label is locked. Close it in Word and try again: $Path"
    }
}

function Invoke-Preflight {
    param([bool]$RequireOcr)
    if ($SkipPreflight) { return }
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $windowsPowerShell)) {
        $windowsPowerShell = (Get-Command powershell.exe -ErrorAction Stop).Source
    }
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $setupScript, '-OcrBackend', $OcrBackend)
    if (-not [string]::IsNullOrWhiteSpace($PythonExe)) { $arguments += @('-PythonExe', $PythonExe) }
    if (-not [string]::IsNullOrWhiteSpace($PandocPath)) { $arguments += @('-PandocPath', $PandocPath) }
    if ($RequireOcr) { $arguments += '-RequireOcr' }
    & $windowsPowerShell @arguments
    if ($LASTEXITCODE -ne 0) {
        throw 'Preflight failed. Install the missing required tools or use -SkipPreflight only after a manual environment check.'
    }
    $script:preflightRuns++
}

function Resolve-PythonCommand {
    param([string]$Configured)
    foreach ($candidate in @($Configured, 'python', 'py', 'python3')) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($null -ne $cmd) {
            if ($candidate -eq 'py') { return @($cmd.Source, '-3') }
            return @($cmd.Source)
        }
    }
    throw 'Python 3 was not found. Run scripts/setup_environment.ps1.'
}

function Invoke-PythonScript {
    param([string[]]$Arguments, [string]$Label)
    $python = @(Resolve-PythonCommand -Configured $PythonExe)
    $exe = $python[0]
    $all = @()
    if ($python.Count -gt 1) { $all += $python[1..($python.Count - 1)] }
    $all += $Arguments
    $output = & $exe @all 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "$Label failed: $($output -join [Environment]::NewLine)"
    }
    return ($output -join [Environment]::NewLine)
}

function Get-PandocPath {
    if (-not [string]::IsNullOrWhiteSpace($PandocPath)) { return $PandocPath }
    $text = Invoke-PythonScript -Arguments @('-c', 'import pypandoc; print(pypandoc.get_pandoc_path())') -Label 'pandoc lookup'
    return $text.Trim()
}

function New-ImportedDocx {
    param([string]$Source, [string]$SourceExtension, [string]$Destination)
    switch ($SourceExtension) {
        { @('.tex','.md','.markdown','.html','.htm','.txt') -contains $_ } {
            $pandoc = Get-PandocPath
            $inputFormat = switch ($SourceExtension) {
                '.tex' { 'latex' }
                { @('.html','.htm') -contains $_ } { 'html' }
                default { 'markdown' }
            }
            & $pandoc $Source '-f' $inputFormat '-t' 'docx' '-o' $Destination
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $Destination)) {
                throw "Pandoc could not import $SourceExtension file: $Source"
            }
            return
        }
        { @('.mml','.xml') -contains $_ } {
            $pandoc = Get-PandocPath
            & $pandoc $Source '-f' 'mathml' '-t' 'docx' '-o' $Destination
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $Destination)) {
                throw "Pandoc could not import MathML file: $Source"
            }
            return
        }
        { @('.png','.jpg','.jpeg','.bmp','.tif','.tiff','.gif','.webp') -contains $_ } {
            $importWord = $null
            $importDoc = $null
            try {
                $importWord = New-Object -ComObject Word.Application
                $importWord.Visible = $false
                $importWord.DisplayAlerts = 0
                $importDoc = $importWord.Documents.Add()
                $picture = $importDoc.InlineShapes.AddPicture($Source, $false, $true)
                try { $picture.AlternativeText = 'formula: imported formula image' } catch {}
                $importDoc.SaveAs2($Destination, $wordFormatDocx)
            } finally {
                if ($importDoc) { try { $importDoc.Close($false) } catch {} }
                if ($importWord) { try { $importWord.Quit() } catch {} }
                [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            }
            return
        }
        default {
            throw "Unsupported source type '$SourceExtension'. Use .docx, .tex, .md, .html, .mml, or an image."
        }
    }
}

function Test-MathTypeReady {
    param($Word)
    if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4')) { return $false }
    foreach ($addIn in $Word.AddIns) {
        try {
            if ($addIn.Installed -and $addIn.Name -like 'MathType Commands*') { return $true }
        } catch {}
    }
    return $false
}

function Find-HeadingStart {
    param($Document, [string]$Heading)
    if ([string]::IsNullOrWhiteSpace($Heading)) { return 0 }
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        if ($paragraph.Range.Text -like ($Heading + '*')) { return [int]$paragraph.Range.Start }
    }
    throw "Section heading not found: $Heading"
}

function Find-SectionEnd {
    param($Document, [int]$Start, [string]$EndHeadingText)
    if (-not [string]::IsNullOrWhiteSpace($EndHeadingText)) {
        for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
            $paragraph = $Document.Paragraphs.Item($i)
            if ($paragraph.Range.Start -le $Start) { continue }
            if ($paragraph.Range.Text -like ($EndHeadingText + '*')) { return [int]$paragraph.Range.Start }
        }
        throw "End heading not found: $EndHeadingText"
    }
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $paragraph = $Document.Paragraphs.Item($i)
        if ($paragraph.Range.Start -le $Start) { continue }
        $text = $paragraph.Range.Text.Trim()
        if ($text -match '^(一|二|三|四|五|六|七|八|九|十)、' -or $text -match '^第[一二三四五六七八九十]+') {
            return [int]$paragraph.Range.Start
        }
    }
    return [int]$Document.Content.End
}

function Get-CleanParagraphText {
    param($Paragraph)
    return ($Paragraph.Range.Text -replace '[\u0001\u0007\u000b\r\n]+$', '' -replace '[\u0001\u0007\u000b]', '')
}

function Test-HasFormulaSignature {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    return ($Text -match '\\[A-Za-z]+' -or $Text -match '[_^][{A-Za-z0-9\\]' -or
            $Text -match '[=<>]' -or $Text -match '\\begin\{' -or $Text -match '\{.*\}')
}

function Get-DelimitedLatexCandidates {
    param($Paragraph)
    $body = Get-CleanParagraphText -Paragraph $Paragraph
    $patterns = @(
        [pscustomobject]@{ Pattern = '\$\$(?<body>[\s\S]+?)\$\$'; Display = $true },
        [pscustomobject]@{ Pattern = '(?<!\$)\$(?!\$)(?<body>[\s\S]+?)(?<!\$)\$(?!\$)'; Display = $false },
        [pscustomobject]@{ Pattern = '\\\[(?<body>[\s\S]+?)\\\]'; Display = $true },
        [pscustomobject]@{ Pattern = '\\\((?<body>[\s\S]+?)\\\)'; Display = $false }
    )
    $raw = New-Object System.Collections.Generic.List[object]
    foreach ($spec in $patterns) {
        foreach ($match in [regex]::Matches($body, $spec.Pattern)) {
            $group = $match.Groups['body']
            $latex = $group.Value.Trim()
            if (-not (Test-HasFormulaSignature -Text $latex)) { continue }
            $leading = $group.Value.Length - $group.Value.TrimStart().Length
            $trimmedLength = $group.Value.Trim().Length
            $start = [int]$Paragraph.Range.Start + $group.Index + $leading
            $raw.Add([pscustomobject]@{
                Kind = 'Latex'; Start = $start; End = $start + $trimmedLength
                Latex = $latex; Display = [bool]$spec.Display; Confidence = 1.0; ShapeNumber = 0
                Source = 'delimited'; Paragraph = $Paragraph
            })
        }
    }
    $ordered = @($raw | Sort-Object Start, @{Expression={ $_.End - $_.Start }; Descending=$true})
    $selected = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in $ordered) {
        $overlap = $false
        foreach ($existing in $selected) {
            if (-not ($candidate.End -le $existing.Start -or $candidate.Start -ge $existing.End)) {
                $overlap = $true; break
            }
        }
        if (-not $overlap) { $selected.Add($candidate) }
    }
    return $selected.ToArray()
}

function Test-ParagraphIntersectsOmml {
    param($Document, $Paragraph)
    for ($i = 1; $i -le $Document.OMaths.Count; $i++) {
        $om = $Document.OMaths.Item($i)
        if ($om.Range.Start -lt $Paragraph.Range.End -and $om.Range.End -gt $Paragraph.Range.Start) { return $true }
    }
    return $false
}

function Get-HeuristicLatexCandidate {
    param($Document, $Paragraph)
    $body = (Get-CleanParagraphText -Paragraph $Paragraph).Trim()
    if ([string]::IsNullOrWhiteSpace($body)) { return $null }
    if (Test-ParagraphIntersectsOmml -Document $Document -Paragraph $Paragraph) { return $null }
    if ($body.Length -gt 800 -or -not (Test-HasFormulaSignature -Text $body)) { return $null }
    if ($body -notmatch '\\[A-Za-z]+|\\begin\{|[_^]\{') { return $null }
    if ($body -match '。$' -and $body -notmatch '\\') { return $null }
    $leading = $Paragraph.Range.Text.Length - $Paragraph.Range.Text.TrimStart().Length
    $start = [int]$Paragraph.Range.Start + $leading
    return [pscustomobject]@{
        Kind = 'Latex'; Start = $start; End = $start + $body.Length
        Latex = $body; Display = $true; Confidence = 0.8; ShapeNumber = 0
        Source = 'heuristic'; Paragraph = $Paragraph
    }
}

function Get-ParagraphLatexCandidates {
    param($Document, [int]$Start, [int]$End, [string]$Mode)
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($paragraph in $Document.Paragraphs) {
        if ($paragraph.Range.End -le $Start -or $paragraph.Range.Start -ge $End) { continue }
        foreach ($candidate in @(Get-DelimitedLatexCandidates -Paragraph $paragraph)) { $items.Add($candidate) }
        if ($Mode -eq 'Heuristic') {
            $candidate = Get-HeuristicLatexCandidate -Document $Document -Paragraph $paragraph
            if ($null -ne $candidate) { $items.Add($candidate) }
        }
    }
    return $items.ToArray()
}

function Get-ImageNumberSet {
    param([string]$Value)
    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    if ([string]::IsNullOrWhiteSpace($Value)) { return ,$set }
    foreach ($part in ($Value -split ',')) {
        $number = 0
        if ([int]::TryParse($part.Trim(), [ref]$number) -and $number -gt 0) { [void]$set.Add($number) }
    }
    return ,$set
}

function Save-InlineShapeImage {
    param($Document, $Shape)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $path = Join-Path $WorkDir ('word_image_' + [guid]::NewGuid().ToString('N') + '.png')
    $Shape.Range.CopyAsPicture()
    $image = $null
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        Start-Sleep -Milliseconds 150
        try { $image = [System.Windows.Forms.Clipboard]::GetImage() } catch {}
        if ($null -ne $image) { break }
    }
    if ($null -eq $image) { throw 'Could not retrieve the picture from the Word clipboard.' }
    try { $image.Save($path, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $image.Dispose() }
    return $path
}

function Test-ImageCandidate {
    param($Shape, [string]$Policy, $NumberSet, [int]$ShapeNumber)
    if ($Policy -eq 'None' -and -not $NumberSet.Contains($ShapeNumber)) { return $false }
    if ($NumberSet.Contains($ShapeNumber)) { return $true }
    if ($Policy -eq 'All') { return $true }
    if ($Policy -eq 'None') { return $false }
    $alt = ''
    try { $alt = [string]$Shape.AlternativeText } catch {}
    $title = ''
    try { $title = [string]$Shape.Title } catch {}
    if (($alt + ' ' + $title) -match 'formula|equation|公式|数学式|math') { return $true }
    if ($Policy -eq 'AltText') { return $false }
    try {
        $paragraph = $Shape.Range.Paragraphs.Item(1)
        $body = (Get-CleanParagraphText -Paragraph $paragraph).Trim()
        if ($body.Length -le 50 -and $body -match '(公式|数学式|equation|formula)') { return $true }
        if ([string]::IsNullOrWhiteSpace($body) -and [double]$Shape.Height -le 300) { return $true }
    } catch {}
    return $false
}

function Get-ImageShapeCandidates {
    param($Document, [int]$Start, [int]$End, [string]$Policy, [string]$Numbers)
    $numberSet = Get-ImageNumberSet -Value $Numbers
    $items = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        if ($shape.Range.Start -lt $Start -or $shape.Range.End -gt $End) { continue }
        $progId = ''
        try { $progId = [string]$shape.OLEFormat.ProgID } catch {}
        if ($progId -eq 'Equation.DSMT4') { continue }
        if (-not (Test-ImageCandidate -Shape $shape -Policy $Policy -NumberSet $numberSet -ShapeNumber $i)) { continue }
        $items.Add([pscustomobject]@{
            Kind = 'Image'; Start = [int]$shape.Range.Start; End = [int]$shape.Range.End
            ShapeNumber = $i; Shape = $shape; Confidence = 0.0; Latex = ''; Display = $false; Source = 'image'
            AltText = $(try { [string]$shape.AlternativeText } catch { '' })
        })
    }
    return $items.ToArray()
}

function Get-ManualOcrResult {
    param($ManualMap, [int]$ShapeNumber)
    if ($null -eq $ManualMap) { return $null }
    if ($ManualMap -is [array]) {
        return $ManualMap | Where-Object { [int]$_.imageNumber -eq $ShapeNumber } | Select-Object -First 1
    }
    $property = $ManualMap.PSObject.Properties | Where-Object { $_.Name -eq [string]$ShapeNumber } | Select-Object -First 1
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Invoke-FormulaOcr {
    param([string]$ImagePath, [int]$ShapeNumber, $ManualMap)
    if ($OcrBackend -eq 'Manual') {
        $manual = Get-ManualOcrResult -ManualMap $ManualMap -ShapeNumber $ShapeNumber
        if ($null -eq $manual) { throw "No manual OCR result for image #$ShapeNumber." }
        return $manual
    }
    $jsonPath = Join-Path $WorkDir ('ocr_' + $ShapeNumber + '_' + [guid]::NewGuid().ToString('N') + '.json')
    $args = @($ocrScript, '--image', $ImagePath, '--output', $jsonPath, '--backend', $OcrBackend.ToLowerInvariant())
    if (-not [string]::IsNullOrWhiteSpace($OcrModel)) { $args += @('--model', $OcrModel) }
    if (-not [string]::IsNullOrWhiteSpace($OcrBaseUrl)) { $args += @('--base-url', $OcrBaseUrl) }
    if (-not [string]::IsNullOrWhiteSpace($OcrApiKey)) { $args += @('--api-key', $OcrApiKey) }
    $null = Invoke-PythonScript -Arguments $args -Label ("OCR for image #$ShapeNumber")
    return (Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json)
}

function Remove-AdjacentDelimiters {
    param($Document, [int]$ShapeStart, [int]$ShapeEnd)
    while ($true) {
        $removed = $false

        if ($ShapeStart -ge 2) {
            $pair = $Document.Range($ShapeStart - 2, $ShapeStart).Text
            if ($pair -eq '$$' -or $pair -eq '\(' -or $pair -eq '\[') {
                $Document.Range($ShapeStart - 2, $ShapeStart).Text = ''
                $ShapeStart -= 2
                $ShapeEnd -= 2
                $removed = $true
            }
        }
        if (-not $removed -and $ShapeStart -ge 1) {
            $single = $Document.Range($ShapeStart - 1, $ShapeStart).Text
            if ($single -eq '$') {
                $Document.Range($ShapeStart - 1, $ShapeStart).Text = ''
                $ShapeStart--
                $ShapeEnd--
                $removed = $true
            }
        }

        if (-not $removed) {
            $contentEnd = [int]$Document.Content.End
            $probeEnd = [Math]::Min($ShapeEnd + 5, $contentEnd)
            if ($probeEnd -gt $ShapeEnd) {
                $probe = $Document.Range($ShapeEnd, $probeEnd).Text
                $match = [regex]::Match($probe, '^[ \t]*(\$\$?|\\\)|\\\])')
                if ($match.Success) {
                    $Document.Range($ShapeEnd, $ShapeEnd + $match.Length).Text = ''
                    $removed = $true
                }
            }
        }

        if (-not $removed) { break }
    }
    return [pscustomobject]@{ Start = $ShapeStart; End = $ShapeEnd }
}

function Find-MathTypeShapeNear {
    param($Document, [int]$Start, [int]$Span)
    $lower = [Math]::Max(0, $Start - 5)
    $upper = $Start + $Span + 5
    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        try {
            if ([string]$shape.OLEFormat.ProgID -eq 'Equation.DSMT4' -and
                $shape.Range.Start -ge $lower -and $shape.Range.Start -le $upper) {
                return $shape
            }
        } catch {}
    }
    return $null
}

function Convert-RangeToMathType {
    param($Word, $Document, [int]$Start, [int]$End, [string]$Latex, [bool]$Display)
    $trimmed = $Latex.Trim()
    # A trailing backslash is usually Pandoc's control-space or an incomplete
    # line break. MathType renders it as a stray glyph, so drop trailing slashes.
    while ($trimmed.EndsWith('\')) {
        $trimmed = $trimmed.Substring(0, $trimmed.Length - 1).TrimEnd()
    }
    # Pandoc emits cases as left-brace + matrix + invisible right delimiter.
    # MathType is more reliable with its native cases environment.
    $trimmed = [regex]::Replace(
        $trimmed,
        '\\left(?:\\\{|\\lbrace)\s*\\begin\{matrix\}([\s\S]*?)\\end\{matrix\}\s*\\right\.',
        '\begin{cases}$1\end{cases}'
    )
    if ($trimmed -match '^[A-Za-z]+$') { $trimmed = '\mathrm{' + $trimmed + '}' }
    $wrapped = if ($Display) { '$$' + $trimmed + '$$' } else { '$' + $trimmed + '$' }
    $beforeShapes = [int]$Document.InlineShapes.Count
    try {
        if ($End -gt $Start) { $Document.Range($Start, $End).Delete() }
        $Document.Range($Start, $Start).Text = $wrapped
    } catch {
        return [pscustomobject]@{ Success=$false; Message=('replace failed: ' + $_.Exception.Message); Start=$Start; End=$Start; Latex=$trimmed }
    }

    $range = $Document.Range($Start, $Start + $wrapped.Length)
    $null = $range.Select()
    $callback = $null
    $message = ''
    try { $null = $Word.Run('MTCommand_OnTexToggle', [ref]$callback) } catch { $message = $_.Exception.Message }
    Start-Sleep -Milliseconds 350
    $elapsed = 350
    $mathShape = Find-MathTypeShapeNear -Document $Document -Start $Start -Span $wrapped.Length
    while ($null -eq $mathShape -and $elapsed -lt 6000) {
        Start-Sleep -Milliseconds 150
        $elapsed += 150
        $mathShape = Find-MathTypeShapeNear -Document $Document -Start $Start -Span $wrapped.Length
    }
    if ($null -ne $mathShape) {
        # Let MathType finish updating the OLE object before editing delimiters.
        Start-Sleep -Milliseconds 350
        $elapsed += 350
        $mathShape = Find-MathTypeShapeNear -Document $Document -Start $Start -Span $wrapped.Length
    }

    if ($null -eq $mathShape) {
        try { $Document.Range($Start, $Start + $wrapped.Length).Text = $trimmed } catch {}
        return [pscustomobject]@{ Success=$false; Message=('MathType toggle timeout: ' + $message); Start=$Start; End=$Start + $wrapped.Length; Latex=$trimmed }
    }

    $shapeStart = [int]$mathShape.Range.Start
    $shapeEnd = [int]$mathShape.Range.End
    return [pscustomobject]@{ Success=$true; Message=$message; Start=$shapeStart; End=$shapeEnd; Latex=$trimmed; ElapsedMs=$elapsed }
}

function Test-Overlap {
    param([int]$StartA, [int]$EndA, [int]$StartB, [int]$EndB)
    return (-not ($EndA -le $StartB -or $StartA -ge $EndB))
}

Invoke-Preflight -RequireOcr $initialRequireOcr

$targetDirectory = Split-Path -Parent $targetPath
if (-not [string]::IsNullOrWhiteSpace($targetDirectory)) {
    New-Item -ItemType Directory -Force -Path $targetDirectory | Out-Null
}
if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $WorkDir = Join-Path ([IO.Path]::GetTempPath()) ('formula_mathtype_' + [guid]::NewGuid().ToString('N'))
} else {
    $WorkDir = [IO.Path]::GetFullPath($WorkDir)
}
New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

$startTime = Get-Date
$events = New-Object System.Collections.Generic.List[object]
$ocrEvents = New-Object System.Collections.Generic.List[object]
$word = $null
$doc = $null
$importedDocx = Join-Path $WorkDir 'imported.docx'
$workingDocx = if ($DryRun) { $sourcePath } else { $targetPath }

try {
    if ($sourceExt -ne '.docx') {
        New-ImportedDocx -Source $sourcePath -SourceExtension $sourceExt -Destination $importedDocx
        if (-not $DryRun) {
            Copy-Item -LiteralPath $importedDocx -Destination $targetPath -Force
            $workingDocx = $targetPath
        } else {
            $workingDocx = $importedDocx
        }
    } else {
        if (-not $DryRun) {
            Assert-FileClosed -Path $targetPath -Label 'Target document'
            Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
            $workingDocx = $targetPath
        }
    }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    $doc = $word.Documents.Open($workingDocx, $false, [bool]$DryRun)

    if (-not $DryRun -and -not (Test-MathTypeReady -Word $word)) {
        throw "MathType's Word add-in is not ready. Install MathType 7, ensure 'MathType Commands 2016.dotm' is loaded, and restart Word."
    }

    $sectionStart = Find-HeadingStart -Document $doc -Heading $SectionHeading
    $sectionEnd = if ([string]::IsNullOrWhiteSpace($SectionHeading)) { [int]$doc.Content.End } else { Find-SectionEnd -Document $doc -Start $sectionStart -EndHeadingText $EndHeading }

    $processOmml = @('Auto','All','Omml') -contains $Mode
    $processLatex = @('Auto','All','Latex') -contains $Mode
    $processImages = @('Auto','All','Images') -contains $Mode
    $candidates = New-Object System.Collections.Generic.List[object]

    if ($processOmml -and $doc.OMaths.Count -gt 0) {
        $mapPath = Join-Path $WorkDir 'omml_math_map.json'
        $mapArgs = @($extractScript, '--source', $workingDocx, '--output', $mapPath)
        if (-not [string]::IsNullOrWhiteSpace($PandocPath)) { $mapArgs += @('--pandoc', $PandocPath) }
        $null = Invoke-PythonScript -Arguments $mapArgs -Label 'OMML extraction'
        $map = Get-Content -LiteralPath $mapPath -Raw -Encoding UTF8 | ConvertFrom-Json
        for ($i = 1; $i -le $doc.OMaths.Count; $i++) {
            $om = $doc.OMaths.Item($i)
            if ($om.Range.Start -lt $sectionStart -or $om.Range.End -gt $sectionEnd) { continue }
            $item = $map | Where-Object { [int]$_.index -eq $i } | Select-Object -First 1
            if ($null -eq $item) { continue }
            $candidates.Add([pscustomobject]@{
                Kind='Omml'; Start=[int]$om.Range.Start; End=[int]$om.Range.End
                Latex=[string]$item.latex; Display=[bool]$item.display; Confidence=1.0; Source='omml'; ShapeNumber=0
            })
        }
    }

    if ($processLatex) {
        foreach ($candidate in @(Get-ParagraphLatexCandidates -Document $doc -Start $sectionStart -End $sectionEnd -Mode $LatexMode)) {
            $overlap = $false
            foreach ($existing in $candidates) {
                if (Test-Overlap -StartA $candidate.Start -EndA $candidate.End -StartB $existing.Start -EndB $existing.End) { $overlap = $true; break }
            }
            if (-not $overlap) { $candidates.Add($candidate) }
        }
    }

    if ($processImages -and $ImagePolicy -ne 'None') {
        $manualMap = $null
        if ($OcrBackend -eq 'Manual') {
            if ([string]::IsNullOrWhiteSpace($ManualOcrMapPath)) { throw 'Manual OCR requires -ManualOcrMapPath.' }
            $manualMap = Get-Content -LiteralPath (Resolve-Path -LiteralPath $ManualOcrMapPath).Path -Raw -Encoding UTF8 | ConvertFrom-Json
        }
        $imageShapes = @(Get-ImageShapeCandidates -Document $doc -Start $sectionStart -End $sectionEnd -Policy $ImagePolicy -Numbers $ImageNumbers)
        if ($imageShapes.Count -gt 0 -and $OcrBackend -ne 'Manual') {
            Invoke-Preflight -RequireOcr $true
        }
        foreach ($image in $imageShapes) {
            $imagePath = $null
            $result = $null
            try {
                $imagePath = Save-InlineShapeImage -Document $doc -Shape $image.Shape
                $result = Invoke-FormulaOcr -ImagePath $imagePath -ShapeNumber $image.ShapeNumber -ManualMap $manualMap
                $ocrEvents.Add([pscustomobject]@{
                    imageNumber=$image.ShapeNumber; altText=$image.AltText; imagePath=$imagePath
                    isFormula=[bool]$result.is_formula; latex=[string]$result.latex
                    confidence=[double]$result.confidence; notes=[string]$result.notes; status='recognized'
                })
                if (-not [bool]$result.is_formula) { continue }
                if ([double]$result.confidence -lt $MinOcrConfidence) { continue }
                if ([string]::IsNullOrWhiteSpace([string]$result.latex)) { continue }
                $image.Latex = [string]$result.latex
                $image.Display = [bool]$result.display
                $image.Confidence = [double]$result.confidence
                $candidates.Add($image)
            } catch {
                $ocrEvents.Add([pscustomobject]@{
                    imageNumber=$image.ShapeNumber; altText=$image.AltText; imagePath=$imagePath
                    isFormula=$false; latex=''; confidence=0.0; notes=$_.Exception.Message; status='failed'
                })
            }
        }
    }

    $ordered = @($candidates | Sort-Object Start)
    if ($MaxFormulas -gt 0) { $ordered = @($ordered | Select-Object -First $MaxFormulas) }
    $ordered = @($ordered | Sort-Object Start -Descending)

    $converted = 0
    $failed = 0
    if (-not $DryRun) {
        foreach ($candidate in $ordered) {
            $resultItems = @(Convert-RangeToMathType -Word $word -Document $doc -Start $candidate.Start -End $candidate.End -Latex $candidate.Latex -Display ([bool]$candidate.Display))
            $result = $resultItems[$resultItems.Count - 1]
            if ($null -eq $result) { $result = [pscustomobject]@{Success=$false;Message='converter returned no result object';ElapsedMs=0} }
            if ($result.Success) {
                $converted++
                $events.Add([pscustomobject]@{kind=$candidate.Kind;source=$candidate.Source;status='converted';start=$candidate.Start;latex=$result.Latex;display=$candidate.Display;confidence=$candidate.Confidence;shapeNumber=$candidate.ShapeNumber;elapsedMs=$result.ElapsedMs})
            } else {
                $failed++
                $events.Add([pscustomobject]@{kind=$candidate.Kind;source=$candidate.Source;status='failed';start=$candidate.Start;latex=$result.Latex;display=$candidate.Display;confidence=$candidate.Confidence;shapeNumber=$candidate.ShapeNumber;message=$result.Message})
            }
            if ($converted -gt 0 -and $converted % 20 -eq 0) { $doc.Save() }
        }
        # Delay delimiter cleanup until every OLE object has finished updating.
        # Cleaning immediately after each conversion can make MathType retain a
        # literal dollar inside some native equation objects.
        Start-Sleep -Milliseconds 500
        # Two passes are intentional: deleting one delimiter shifts the next
        # OLE object's range, and the final object is often the affected one.
        for ($pass = 1; $pass -le 2; $pass++) {
            for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
                $shape = $doc.InlineShapes.Item($i)
                try {
                    if ([string]$shape.OLEFormat.ProgID -eq 'Equation.DSMT4') {
                        $null = Remove-AdjacentDelimiters -Document $doc -ShapeStart ([int]$shape.Range.Start) -ShapeEnd ([int]$shape.Range.End)
                    }
                } catch {}
            }
        }
        $doc.Save()
    } else {
        foreach ($candidate in $ordered) {
            $events.Add([pscustomobject]@{kind=$candidate.Kind;source=$candidate.Source;status='planned';start=$candidate.Start;latex=$candidate.Latex;display=$candidate.Display;confidence=$candidate.Confidence;shapeNumber=$candidate.ShapeNumber})
        }
    }

    if (-not $DryRun) {
        if ([string]::IsNullOrWhiteSpace($SectionHeading)) {
            $sectionStart = 0
            $sectionEnd = [int]$doc.Content.End
        } else {
            $sectionStart = Find-HeadingStart -Document $doc -Heading $SectionHeading
            $sectionEnd = Find-SectionEnd -Document $doc -Start $sectionStart -EndHeadingText $EndHeading
        }
    }

    $remainingOmml = 0
    for ($i = 1; $i -le $doc.OMaths.Count; $i++) {
        $om = $doc.OMaths.Item($i)
        if ($om.Range.Start -ge $sectionStart -and $om.Range.End -le $sectionEnd) { $remainingOmml++ }
    }
    $mathTypeShapes = 0
    $otherShapes = 0
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $shape = $doc.InlineShapes.Item($i)
        if ($shape.Range.Start -lt $sectionStart -or $shape.Range.End -gt $sectionEnd) { continue }
        $progId = ''
        try { $progId = [string]$shape.OLEFormat.ProgID } catch {}
        if ($progId -eq 'Equation.DSMT4') { $mathTypeShapes++ } else { $otherShapes++ }
    }
    $sectionText = $doc.Range($sectionStart, $sectionEnd).Text
    $dollarCount = ([regex]::Matches($sectionText, '\$')).Count

    $summary = [pscustomobject]@{
        source = $sourcePath
        target = if ($DryRun) { '' } else { $targetPath }
        mode = $Mode
        latexMode = $LatexMode
        imagePolicy = $ImagePolicy
        dryRun = [bool]$DryRun
        preflightSkipped = [bool]$SkipPreflight
        preflightRuns = $preflightRuns
        sectionHeading = $SectionHeading
        selectedCandidates = $candidates.Count
        processed = $ordered.Count
        ommlCandidates = @($candidates | Where-Object Kind -eq 'Omml').Count
        latexCandidates = @($candidates | Where-Object Kind -eq 'Latex').Count
        imageCandidates = @($candidates | Where-Object Kind -eq 'Image').Count
        converted = $converted
        failed = $failed
        remainingOMaths = $remainingOmml
        mathTypeShapes = $mathTypeShapes
        otherInlineShapes = $otherShapes
        remainingDollarCharacters = $dollarCount
        elapsedSeconds = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
    }
    $logPath = if ($DryRun) { $targetPath + '.dryrun.json' } else { $targetPath + '.log.json' }
    $log = [pscustomobject]@{ summary=$summary; ocr=$ocrEvents.ToArray(); events=$events.ToArray() }
    $log | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $logPath -Encoding UTF8
    $summary | Format-List
    "Log: $logPath"
    if ($failed -gt 0) { exit 2 }
} finally {
    if ($doc) { try { $doc.Close($false) } catch {} }
    if ($word) { try { $word.Quit() } catch {} }
    try { if ($word) { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } } catch {}
    $word = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if (-not $KeepWorkFiles) { try { Remove-Item -LiteralPath $WorkDir -Recurse -Force } catch {} }
}
