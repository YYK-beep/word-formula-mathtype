param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [string]$TargetPath,
    [string]$MapPath,
    [string]$SectionHeading,
    [string]$EndHeading,
    [int]$MaxFormulas = 0,
    [string]$PandocPath,
    [string]$PythonExe = "python",
    [switch]$Force,
    [switch]$SkipPreflight
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$extractScript = Join-Path $scriptDir 'extract_math_map.py'
$setupScript = Join-Path $scriptDir 'setup_environment.ps1'

function Assert-FileClosed {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
    } catch {
        throw "$Label is locked. Close it in Word and try again: $Path"
    }
}

function Invoke-Preflight {
    if ($SkipPreflight) { return }
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $windowsPowerShell)) {
        $windowsPowerShell = (Get-Command powershell.exe -ErrorAction Stop).Source
    }
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $setupScript, '-OcrBackend', 'None')
    if (-not [string]::IsNullOrWhiteSpace($PythonExe)) { $arguments += @('-PythonExe', $PythonExe) }
    if (-not [string]::IsNullOrWhiteSpace($PandocPath)) { $arguments += @('-PandocPath', $PandocPath) }
    & $windowsPowerShell @arguments
    if ($LASTEXITCODE -ne 0) {
        throw 'Preflight failed. Install the missing required tools or use -SkipPreflight only after a manual environment check.'
    }
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
    throw "Python was not found. Install Python 3.10+ and run scripts/setup_environment.ps1."
}

function Invoke-MapExtraction {
    param([string]$Source, [string]$Output)
    $python = @(Resolve-PythonCommand -Configured $PythonExe)
    $exe = $python[0]
    $args = @()
    if ($python.Count -gt 1) { $args += $python[1..($python.Count - 1)] }
    $args += @($extractScript, '--source', $Source, '--output', $Output)
    if (-not [string]::IsNullOrWhiteSpace($PandocPath)) { $args += @('--pandoc', $PandocPath) }
    & $exe @args | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "Math map extraction failed." }
}

function Find-HeadingStart {
    param($Document, [string]$Heading)
    if ([string]::IsNullOrWhiteSpace($Heading)) { return 0 }
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $t = $Document.Paragraphs.Item($i).Range.Text
        if ($t -like ($Heading + '*')) { return $Document.Paragraphs.Item($i).Range.Start }
    }
    throw "Section heading not found: $Heading"
}

function Find-SectionEnd {
    param($Document, [int]$Start, [string]$EndHeadingText)
    if (-not [string]::IsNullOrWhiteSpace($EndHeadingText)) {
        for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
            $p = $Document.Paragraphs.Item($i)
            if ($p.Range.Start -le $Start) { continue }
            if ($p.Range.Text -like ($EndHeadingText + '*')) { return $p.Range.Start }
        }
        throw "End heading not found: $EndHeadingText"
    }
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $p = $Document.Paragraphs.Item($i)
        if ($p.Range.Start -le $Start) { continue }
        $t = $p.Range.Text.Trim()
        if ($t -match '^(一|二|三|四|五|六|七|八|九|十)、' -or $t -match '^第[一二三四五六七八九十]+') {
            return $p.Range.Start
        }
    }
    return $Document.Content.End
}

function Find-InlineShapeNearStart {
    param($Document, [int]$Start, [int]$MaxStart)
    for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
        $shape = $Document.InlineShapes.Item($i)
        if ($shape.Range.Start -ge $Start -and $shape.Range.Start -le $MaxStart) { return $shape }
    }
    return $null
}

function Remove-AdjacentDollars {
    param($Document, $Shape)
    while ($true) {
        $removed = $false
        $shapeStart = [int]$Shape.Range.Start
        $shapeEnd = [int]$Shape.Range.End
        $contentEnd = [int]$Document.Content.End
        if ($shapeStart -gt 0) {
            $lead = $Document.Range($shapeStart - 1, $shapeStart)
            if ($lead.Text -eq '$') { $lead.Text = ''; $removed = $true }
        }
        if ($shapeEnd -lt $contentEnd) {
            $trail = $Document.Range($shapeEnd, [Math]::Min($shapeEnd + 1, $contentEnd))
            if ($trail.Text -eq '$') { $trail.Text = ''; $removed = $true }
        }
        if (-not $removed) { break }
    }
}

if ([string]::IsNullOrWhiteSpace($TargetPath)) {
    $dir = Split-Path -Parent $SourcePath
    $name = [IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $TargetPath = Join-Path $dir ($name + '_mathtype.docx')
}
$SourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
$TargetPath = [IO.Path]::GetFullPath($TargetPath)
if ($SourcePath -eq $TargetPath) { throw "TargetPath must differ from SourcePath." }
if ((Test-Path -LiteralPath $TargetPath) -and -not $Force) {
    throw "Target exists. Pass -Force to overwrite: $TargetPath"
}
Assert-FileClosed -Path $TargetPath -Label 'Target document'
Invoke-Preflight

if ([string]::IsNullOrWhiteSpace($MapPath)) {
    $MapPath = Join-Path ([IO.Path]::GetTempPath()) ("omml_math_map_" + [guid]::NewGuid().ToString('N') + '.json')
}
if (-not (Test-Path -LiteralPath $MapPath) -or -not [string]::IsNullOrWhiteSpace($PandocPath)) {
    Invoke-MapExtraction -Source $SourcePath -Output $MapPath
}

Copy-Item -LiteralPath $SourcePath -Destination $TargetPath -Force

$word = $null
$doc = $null
$events = New-Object System.Collections.Generic.List[object]
$converted = 0
$failed = 0
$skipped = 0
$startTime = Get-Date
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false
    $doc = $word.Documents.Open($TargetPath, $false, $false)

    $mathTypeAddIn = $false
    foreach ($addin in $word.AddIns) {
        if ($addin.Name -eq 'MathType Commands 2016.dotm' -and $addin.Installed) { $mathTypeAddIn = $true }
    }
    if (-not $mathTypeAddIn) {
        throw "MathType's Word add-in was not found. Install MathType 7 and restart Word."
    }

    $sectionStart = Find-HeadingStart -Document $doc -Heading $SectionHeading
    $sectionEnd = if ([string]::IsNullOrWhiteSpace($SectionHeading)) { $doc.Content.End } else { Find-SectionEnd -Document $doc -Start $sectionStart -EndHeadingText $EndHeading }

    $selected = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $doc.OMaths.Count; $i++) {
        $om = $doc.OMaths.Item($i)
        if ($om.Range.Start -ge $sectionStart -and $om.Range.End -le $sectionEnd) {
            $selected.Add([pscustomobject]@{ OriginalIndex = $i })
        }
    }
    if ($selected.Count -eq 0) { throw "No OMath equations found in the selected range." }

    $ordered = @($selected | Sort-Object OriginalIndex -Descending)
    if ($MaxFormulas -gt 0) { $ordered = @($ordered | Select-Object -First $MaxFormulas) }

    $map = Get-Content -LiteralPath $MapPath -Raw -Encoding UTF8 | ConvertFrom-Json

    foreach ($entry in $ordered) {
        $index = [int]$entry.OriginalIndex
        $item = $map | Where-Object { [int]$_.index -eq $index } | Select-Object -First 1
        if ($null -eq $item) { $failed++; $events.Add([pscustomobject]@{index=$index;status='failed';message='No LaTeX mapping'}); continue }

        $om = $null
        try { $om = $doc.OMaths.Item($index) } catch { $failed++; $events.Add([pscustomobject]@{index=$index;status='failed';message='OMath missing'}); continue }
        $start = $om.Range.Start
        $end = $om.Range.End
        $latex = ([string]$item.latex).Trim()
        while ($latex.EndsWith('\')) {
            $latex = $latex.Substring(0, $latex.Length - 1).TrimEnd()
        }
        $display = [bool]$item.display
        if ($latex -match '^[A-Za-z]+$') {
            $wrap = '$\mathrm{' + $latex + '}$'
        } elseif ($display) {
            $wrap = '$$' + $latex + '$$'
        } else {
            $wrap = '$' + $latex + '$'
        }

        try {
            $doc.Range($start, $end).Delete()
            $doc.Range($start, $start).Text = $wrap
        } catch {
            $failed++; $events.Add([pscustomobject]@{index=$index;status='failed';message=('replace: ' + $_.Exception.Message);latex=$latex}); continue
        }

        $work = $doc.Range($start, $start + $wrap.Length)
        $work.Select()
        $beforeShapes = $doc.InlineShapes.Count
        $callback = $null
        $err = ''
        try { $word.Run('MTCommand_OnTexToggle', [ref]$callback) | Out-Null } catch { $err = $_.Exception.Message }
        $elapsed = 0
        $ok = $false
        while ($elapsed -lt 3000) {
            Start-Sleep -Milliseconds 100
            $elapsed += 100
            if ($doc.InlineShapes.Count -gt $beforeShapes) { $ok = $true; break }
        }

        if ($ok) {
            $converted++
            $events.Add([pscustomobject]@{index=$index;status='converted';display=$display;latex=$latex;elapsedMs=$elapsed})
        } else {
            $failed++
            # Keep failed formulas as readable plain LaTeX text rather than raw $...$ noise.
            try { $doc.Range($start, $start + $wrap.Length).Text = $latex } catch {}
            $events.Add([pscustomobject]@{index=$index;status='failed';message=('toggle timeout: ' + $err);latex=$latex})
        }

        if ($converted -gt 0 -and $converted % 20 -eq 0) { $doc.Save() }
    }

    # Final cleanup: remove delimiter dollars adjacent to any MathType object.
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $shape = $doc.InlineShapes.Item($i)
        try {
            if ($shape.OLEFormat.ProgID -eq 'Equation.DSMT4') {
                if ($shape.Range.Start -ge $sectionStart -and $shape.Range.End -le $sectionEnd) {
                    Remove-AdjacentDollars -Document $doc -Shape $shape
                }
            }
        } catch {}
    }

    $doc.Save()
    # Formula replacement changes document length; recompute the selected range
    # before the final measurements instead of reusing stale Word Range offsets.
    if ([string]::IsNullOrWhiteSpace($SectionHeading)) {
        $sectionStart = 0
        $sectionEnd = [int]$doc.Content.End
    } else {
        $sectionStart = Find-HeadingStart -Document $doc -Heading $SectionHeading
        $sectionEnd = Find-SectionEnd -Document $doc -Start $sectionStart -EndHeadingText $EndHeading
    }
    $remainingOMaths = 0
    for ($i = 1; $i -le $doc.OMaths.Count; $i++) {
        $om = $doc.OMaths.Item($i)
        if ($om.Range.Start -ge $sectionStart -and $om.Range.End -le $sectionEnd) { $remainingOMaths++ }
    }
    $mathTypeShapes = 0
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $shape = $doc.InlineShapes.Item($i)
        try {
            if ($shape.OLEFormat.ProgID -eq 'Equation.DSMT4' -and $shape.Range.Start -ge $sectionStart -and $shape.Range.End -le $sectionEnd) { $mathTypeShapes++ }
        } catch {}
    }
    $sectionText = $doc.Range($sectionStart, $sectionEnd).Text
    $dollarCount = ([regex]::Matches($sectionText, '\$')).Count

    $summary = [pscustomobject]@{
        source = $SourcePath
        target = $TargetPath
        map = $MapPath
        sectionHeading = $SectionHeading
        selectedOMaths = $selected.Count
        processed = $ordered.Count
        converted = $converted
        failed = $failed
        remainingOMaths = $remainingOMaths
        mathTypeShapes = $mathTypeShapes
        remainingDollarCharacters = $dollarCount
        elapsedSeconds = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
    }
    $log = [pscustomobject]@{ summary = $summary; events = $events.ToArray() }
    $logPath = $TargetPath + '.log.json'
    $log | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $logPath -Encoding UTF8
    $summary | Format-List
    "Log: $logPath"
} finally {
    if ($doc) { try { $doc.Close($false) } catch {} }
    if ($word) { try { $word.Quit() } catch {} }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    $word = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}


