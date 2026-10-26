param(
    [switch]$InstallMissing = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-CommandPath {
    param([string[]]$Names)
    foreach ($name in $Names) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source)) { return $cmd.Source }
    }
    return $null
}

function Resolve-WordPath {
    $fromPath = Get-CommandPath -Names @('WINWORD.EXE', 'winword')
    if ($fromPath) { return $fromPath }
    $appPath = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE' -ErrorAction SilentlyContinue
    if ($null -ne $appPath -and -not [string]::IsNullOrWhiteSpace($appPath.'(default)')) {
        return $appPath.'(default)'
    }
    foreach ($candidate in @(
        'C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE',
        'C:\Program Files (x86)\Microsoft Office\Root\Office16\WINWORD.EXE'
    )) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    return $null
}

$result = [ordered]@{}
$result.word = Resolve-WordPath
$result.python = Get-CommandPath -Names @('python', 'py', 'python3')

$result.mathTypeRegistered = Test-Path 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4'
$startupCandidates = @(
    'C:\Program Files\Microsoft Office\root\Office16\STARTUP\MathType Commands 2016.dotm',
    'C:\Program Files (x86)\Microsoft Office\root\Office16\STARTUP\MathType Commands 2016.dotm',
    (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\MathType Commands 2016.dotm')
)
$result.mathTypeWordAddIn = ($startupCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)

$pythonExe = $result.python
$result.pypandoc = $false
$result.visionOcrConfigured = [bool](
    $env:FORMULA_OCR_API_KEY -or $env:SN_VISION_API_KEY -or $env:SN_CHAT_API_KEY -or $env:SN_API_KEY -or
    ($env:MATHPIX_APP_ID -and $env:MATHPIX_APP_KEY)
)
$result.defaultOcrModel = if ($env:FORMULA_OCR_MODEL) { $env:FORMULA_OCR_MODEL } elseif ($env:SN_VISION_MODEL) { $env:SN_VISION_MODEL } else { 'sensenova-6.8-flash-lite' }

if ($pythonExe) {
    & $pythonExe -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('pypandoc') else 1)" 1>$null 2>$null
    $result.pypandoc = ($LASTEXITCODE -eq 0)
    if (-not $result.pypandoc -and $InstallMissing) {
        Write-Host 'Installing pypandoc-binary for the current user...'
        & $pythonExe -m pip install --user pypandoc-binary
        if ($LASTEXITCODE -ne 0) { throw 'Failed to install pypandoc-binary.' }
        $result.pypandoc = $true
    }
}

[pscustomobject]$result | Format-List

if (-not $result.word) { Write-Warning 'Microsoft Word was not found. Install Word desktop.' }
if (-not $result.mathTypeRegistered) { Write-Warning 'Equation.DSMT4 is not registered. Install MathType 7 first.' }
if (-not $result.mathTypeWordAddIn) { Write-Warning 'MathType Commands 2016.dotm was not found in Word STARTUP.' }
if (-not $result.python) { Write-Warning 'Python 3 was not found on PATH.' }
if (-not $result.pypandoc) { Write-Warning 'pypandoc-binary is missing. Run this script again with -InstallMissing.' }
