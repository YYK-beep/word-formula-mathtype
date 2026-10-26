param(
    [Parameter(Mandatory=$true)][string]$Path,
    [string]$SectionHeading,
    [string]$EndHeading
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Find-HeadingStart {
    param($Document, [string]$Heading)
    if ([string]::IsNullOrWhiteSpace($Heading)) { return 0 }
    for ($i = 1; $i -le $Document.Paragraphs.Count; $i++) {
        $p = $Document.Paragraphs.Item($i)
        if ($p.Range.Text -like ($Heading + '*')) { return $p.Range.Start }
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
        if ($t -match '^(一|二|三|四|五|六|七|八|九|十)、' -or $t -match '^第[一二三四五六七八九十]+') { return $p.Range.Start }
    }
    return $Document.Content.End
}

function Test-HasFormulaSignature {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    return ($Text -match '\\[A-Za-z]+' -or $Text -match '[_^][{A-Za-z0-9\\]' -or
            $Text -match '[=<>]' -or $Text -match '\\begin\{' -or $Text -match '\{.*\}')
}

function Get-CleanParagraphText {
    param($Paragraph)
    return ($Paragraph.Range.Text -replace '[\u0001\u0007\u000b\r\n]+$', '' -replace '[\u0001\u0007\u000b]', '')
}

function Get-RemainingFormulaTextCounts {
    param($Document, [int]$Start, [int]$End)
    $delimiterCount = 0
    $rawLatexParagraphs = 0
    $patterns = @(
        '\$\$(?<body>[\s\S]+?)\$\$',
        '(?<!\$)\$(?!\$)(?<body>[\s\S]+?)(?<!\$)\$(?!\$)',
        '\\\[(?<body>[\s\S]+?)\\\]',
        '\\\((?<body>[\s\S]+?)\\\)'
    )
    $rawCommand = '\\(?:frac|sqrt|sum|prod|int|iint|iiint|lim|begin|alpha|beta|gamma|delta|nabla|partial|mathrm|mathbf|mathcal|left|right|text)\b'
    foreach ($paragraph in $Document.Paragraphs) {
        if ($paragraph.Range.End -le $Start -or $paragraph.Range.Start -ge $End) { continue }
        $body = Get-CleanParagraphText -Paragraph $paragraph
        $hasDelimitedFormula = $false
        foreach ($pattern in $patterns) {
            foreach ($match in [regex]::Matches($body, $pattern)) {
                if (Test-HasFormulaSignature -Text $match.Groups['body'].Value) {
                    $delimiterCount++
                    $hasDelimitedFormula = $true
                }
            }
        }
        if (-not $hasDelimitedFormula -and $body -match $rawCommand -and (Test-HasFormulaSignature -Text $body)) {
            $rawLatexParagraphs++
        }
    }
    return [pscustomobject]@{ DelimiterCount = $delimiterCount; RawLatexParagraphs = $rawLatexParagraphs }
}

$path = (Resolve-Path -LiteralPath $Path).Path
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($path, $false, $true)

    $start = Find-HeadingStart -Document $doc -Heading $SectionHeading
    $end = if ([string]::IsNullOrWhiteSpace($SectionHeading)) { $doc.Content.End } else { Find-SectionEnd -Document $doc -Start $start -EndHeadingText $EndHeading }

    $omaths = 0
    for ($i = 1; $i -le $doc.OMaths.Count; $i++) {
        $om = $doc.OMaths.Item($i)
        if ($om.Range.Start -ge $start -and $om.Range.End -le $end) { $omaths++ }
    }

    $mathTypeShapes = 0
    $otherShapes = 0
    for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
        $shape = $doc.InlineShapes.Item($i)
        if ($shape.Range.Start -lt $start -or $shape.Range.End -gt $end) { continue }
        $progId = ''
        try { $progId = $shape.OLEFormat.ProgID } catch {}
        if ($progId -eq 'Equation.DSMT4') { $mathTypeShapes++ } else { $otherShapes++ }
    }

    $text = $doc.Range($start, $end).Text
    $dollarCount = ([regex]::Matches($text, '\$')).Count
    $remaining = Get-RemainingFormulaTextCounts -Document $doc -Start $start -End $end
    $result = [pscustomobject]@{
        path = $path
        sectionHeading = $SectionHeading
        omathsInRange = $omaths
        mathTypeShapes = $mathTypeShapes
        otherInlineShapes = $otherShapes
        remainingFormulaDelimiters = $remaining.DelimiterCount
        remainingRawLatexParagraphs = $remaining.RawLatexParagraphs
        remainingDollarCharacters = $dollarCount
        isMathTypeOnly = ($omaths -eq 0 -and $remaining.DelimiterCount -eq 0 -and $remaining.RawLatexParagraphs -eq 0 -and $mathTypeShapes -gt 0)
    }
    $result | Format-List
} finally {
    if ($doc) { try { $doc.Close($false) } catch {} }
    if ($word) { try { $word.Quit() } catch {} }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
    $word = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
