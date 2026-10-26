[CmdletBinding()]
param(
    [switch]$InstallMissing,
    [switch]$RequireOcr,
    [ValidateSet('None','Vlm','Mathpix','Manual')][string]$OcrBackend = 'Vlm',
    [string]$PandocPath,
    [string]$PythonExe,
    [switch]$Json
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$checks = New-Object System.Collections.Generic.List[object]

function Add-Check {
    param(
        [string]$Id,
        [string]$Name,
        [bool]$Required,
        [bool]$Passed,
        [string]$Detail
    )
    $checks.Add([pscustomobject]@{
        id = $Id
        name = $Name
        required = $Required
        passed = $Passed
        status = if ($Passed) { 'PASS' } elseif ($Required) { 'FAIL' } else { 'OPTIONAL' }
        detail = $Detail
    })
}

function Get-CommandPath {
    param([string[]]$Names)
    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
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

function Get-MathTypeUninstallEntry {
    $entries = @()
    foreach ($key in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )) {
        $entries += @(Get-ItemProperty $key -ErrorAction SilentlyContinue | Where-Object { $null -ne $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -match '^MathType\s*7' })
    }
    return ($entries | Select-Object -First 1)
}

function Resolve-MathTypeExecutable {
    $defaultIcon = $null
    try { $defaultIcon = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4\DefaultIcon' -ErrorAction Stop).'(default)' } catch {}
    if (-not [string]::IsNullOrWhiteSpace($defaultIcon)) {
        $match = [regex]::Match($defaultIcon, '^\s*"?([^",]+\.exe)', 'IgnoreCase')
        if ($match.Success -and (Test-Path -LiteralPath $match.Groups[1].Value)) { return $match.Groups[1].Value }
    }
    foreach ($candidate in @('D:\APP\MathType.exe','C:\Program Files\MathType\MathType.exe','C:\Program Files (x86)\MathType\MathType.exe')) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    return $null
}

function Resolve-MathTypeAddIn {
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($candidate in @(
        'C:\Program Files\Microsoft Office\root\Office16\STARTUP\MathType Commands 2016.dotm',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\STARTUP\MathType Commands 2016.dotm',
        (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\MathType Commands 2016.dotm')
    )) {
        $candidates.Add($candidate)
    }
    foreach ($officeRoot in @('C:\Program Files\Microsoft Office','C:\Program Files (x86)\Microsoft Office')) {
        if (-not (Test-Path -LiteralPath $officeRoot)) { continue }
        foreach ($file in @(Get-ChildItem -LiteralPath $officeRoot -Recurse -File -Filter 'MathType Commands*.dotm' -ErrorAction SilentlyContinue)) {
            $candidates.Add($file.FullName)
        }
    }
    return ($candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
}

function Invoke-Python {
    param([string]$Executable, [string[]]$PrefixArgs, [string[]]$Arguments)
    $all = @($PrefixArgs) + @($Arguments)
    return (& $Executable @all 2>&1 | Out-String).Trim()
}

# 1. PowerShell
$psVersion = $PSVersionTable.PSVersion
$psOk = $psVersion -ge [Version]'5.1'
Add-Check -Id 'powershell' -Name 'Windows PowerShell' -Required $true -Passed $psOk -Detail $psVersion.ToString()

# 2. Microsoft Word
$wordPath = Resolve-WordPath
$wordVersion = ''
$wordVersionOk = $false
if ($wordPath) {
    try {
        $wordVersion = (Get-Item -LiteralPath $wordPath).VersionInfo.ProductVersion
        $wordVersionOk = ([version]($wordVersion.Split(' ')[0])).Major -ge 16
    } catch {}
}
Add-Check -Id 'word' -Name 'Microsoft Word desktop 2016+' -Required $true -Passed ([bool]$wordPath -and $wordVersionOk) -Detail $(if ($wordPath) { "$wordPath ($wordVersion)" } else { 'Not found. Install Word 2016 or later.' })

# 3. MathType registration and Word add-in
$mathTypeRegistered = Test-Path 'Registry::HKEY_CLASSES_ROOT\Equation.DSMT4'
$mathTypeExe = Resolve-MathTypeExecutable
$mathTypeInstall = Get-MathTypeUninstallEntry
$mathTypeVersion = if ($null -ne $mathTypeInstall -and $mathTypeInstall.DisplayVersion) { $mathTypeInstall.DisplayVersion } else { '' }
if (-not $mathTypeVersion -and $mathTypeExe) {
    try { $mathTypeVersion = (Get-Item -LiteralPath $mathTypeExe).VersionInfo.ProductVersion } catch {}
}
$mathTypeVersionOk = $mathTypeVersion -match '^7\.'
$mathTypeDetail = if ($mathTypeRegistered) { "Equation.DSMT4 registered" } else { 'Equation.DSMT4 is not registered.' }
if ($mathTypeVersion) { $mathTypeDetail += "; version $mathTypeVersion" }
if ($mathTypeExe) { $mathTypeDetail += "; $mathTypeExe" }
Add-Check -Id 'mathtype-registration' -Name 'MathType 7 OLE registration' -Required $true -Passed $mathTypeRegistered -Detail $mathTypeDetail
Add-Check -Id 'mathtype-version' -Name 'MathType 7.x version' -Required $true -Passed $mathTypeVersionOk -Detail $(if ($mathTypeVersion) { $mathTypeVersion } else { 'MathType version could not be detected.' })

$mathTypeAddIn = Resolve-MathTypeAddIn
Add-Check -Id 'mathtype-word-addin' -Name 'MathType Word add-in' -Required $true -Passed ([bool]$mathTypeAddIn) -Detail $(if ($mathTypeAddIn) { $mathTypeAddIn } else { 'MathType Commands 2016.dotm was not found in Word STARTUP folders.' })

# 4. Python
$configuredPython = if (-not [string]::IsNullOrWhiteSpace($PythonExe)) { Get-CommandPath -Names @($PythonExe) } else { $null }
$pythonPath = if ($configuredPython) { $configuredPython } else { Get-CommandPath -Names @('python', 'python3', 'py') }
$pythonPrefix = @()
if ($pythonPath -and ([IO.Path]::GetFileNameWithoutExtension($pythonPath) -eq 'py')) { $pythonPrefix = @('-3') }
$pythonVersion = ''
$pythonOk = $false
if ($pythonPath) {
    try {
        $pythonVersion = Invoke-Python -Executable $pythonPath -PrefixArgs $pythonPrefix -Arguments @('-c', 'import sys; print(str(sys.version_info[0])+chr(46)+str(sys.version_info[1])+chr(46)+str(sys.version_info[2]))')
        $parts = $pythonVersion.Split('.')
        $pythonOk = ([int]$parts[0] -gt 3) -or ([int]$parts[0] -eq 3 -and [int]$parts[1] -ge 10)
    } catch {
        $pythonVersion = $_.Exception.Message
    }
}
Add-Check -Id 'python' -Name 'Python 3.10+' -Required $true -Passed $pythonOk -Detail $(if ($pythonPath) { "$pythonPath ($pythonVersion)" } else { 'Python was not found on PATH.' })

# 5. Pandoc / pypandoc
$pypandocInstalled = $false
if ($pythonOk) {
    try {
        $probe = Invoke-Python -Executable $pythonPath -PrefixArgs $pythonPrefix -Arguments @('-c', 'import pypandoc; print(1)')
        $pypandocInstalled = ($probe -match '1')
    } catch {}
}
if (-not $pypandocInstalled -and $InstallMissing -and $pythonOk) {
    Write-Host 'Installing pypandoc-binary for the current user...'
    $installOutput = Invoke-Python -Executable $pythonPath -PrefixArgs $pythonPrefix -Arguments @('-m', 'pip', 'install', '--user', 'pypandoc-binary')
    if ($LASTEXITCODE -ne 0) { Write-Warning $installOutput }
    $probe = Invoke-Python -Executable $pythonPath -PrefixArgs $pythonPrefix -Arguments @('-c', 'import pypandoc; print(1)')
    $pypandocInstalled = ($probe -match '1')
}

$pandocExe = $null
if (-not [string]::IsNullOrWhiteSpace($PandocPath)) {
    if (Test-Path -LiteralPath $PandocPath) { $pandocExe = $PandocPath }
} else {
    $pandocExe = Get-CommandPath -Names @('pandoc')
}
if (-not $pandocExe -and $pypandocInstalled -and $pythonOk) {
    try {
        $candidate = Invoke-Python -Executable $pythonPath -PrefixArgs $pythonPrefix -Arguments @('-c', 'import pypandoc; print(pypandoc.get_pandoc_path())')
        if (-not (Test-Path -LiteralPath $candidate) -and (Test-Path -LiteralPath ($candidate + '.exe'))) {
            $candidate = $candidate + '.exe'
        }
        if (Test-Path -LiteralPath $candidate) { $pandocExe = $candidate }
    } catch {}
}
$pandocVersion = ''
$pandocOk = $false
if ($pandocExe) {
    try {
        $firstLine = (& $pandocExe --version 2>&1 | Select-Object -First 1)
        $pandocVersion = [string]$firstLine
        $match = [regex]::Match($pandocVersion, '(\d+)\.(\d+)')
        if ($match.Success) {
            $major = [int]$match.Groups[1].Value
            $pandocOk = ($major -ge 3)
        }
    } catch {}
}
Add-Check -Id 'pandoc' -Name 'Pandoc 3.x / pypandoc-binary' -Required $true -Passed $pandocOk -Detail $(if ($pandocExe) { "$pandocExe ($pandocVersion)" } elseif ($pypandocInstalled) { 'pypandoc installed, but its Pandoc executable was not found.' } else { 'Install pypandoc-binary or Pandoc 3.x.' })

# 6. OCR configuration (required only for image workflows or explicitly requested)
$vlmConfigured = [bool]($env:FORMULA_OCR_API_KEY -or $env:SN_VISION_API_KEY -or $env:SN_CHAT_API_KEY -or $env:SN_API_KEY)
$mathpixConfigured = [bool]($env:MATHPIX_APP_ID -and $env:MATHPIX_APP_KEY)
$ocrReady = switch ($OcrBackend) {
    'None' { $true }
    'Manual' { $true }
    'Mathpix' { $mathpixConfigured }
    default { $vlmConfigured }
}
$ocrDetail = switch ($OcrBackend) {
    'Manual' { 'Manual OCR map selected; API credentials are not required.' }
    'Mathpix' { if ($mathpixConfigured) { 'MATHPIX_APP_ID and MATHPIX_APP_KEY are configured.' } else { 'Set MATHPIX_APP_ID and MATHPIX_APP_KEY.' } }
    'None' { 'OCR is disabled for this run.' }
    default { if ($vlmConfigured) { 'Vision API key is configured.' } else { 'Set FORMULA_OCR_API_KEY or SN_API_KEY when image OCR is needed.' } }
}
Add-Check -Id 'ocr' -Name "OCR backend: $OcrBackend" -Required ([bool]$RequireOcr) -Passed $ocrReady -Detail $ocrDetail

$failedRequired = @($checks | Where-Object { $_.required -and -not $_.passed })
$summary = [pscustomobject]@{
    status = if ($failedRequired.Count -eq 0) { 'PASS' } else { 'FAIL' }
    checkedAt = (Get-Date).ToString('s')
    os = [Environment]::OSVersion.VersionString
    powershell = $psVersion.ToString()
    word = $wordPath
    wordVersion = $wordVersion
    mathTypeVersion = $mathTypeVersion
    mathTypeAddIn = $mathTypeAddIn
    python = $pythonPath
    pythonVersion = $pythonVersion
    pandoc = $pandocExe
    pandocVersion = $pandocVersion
    pypandocInstalled = $pypandocInstalled
    ocrBackend = $OcrBackend
    ocrConfigured = $ocrReady
    missingRequired = @($failedRequired | ForEach-Object { $_.id })
    checks = $checks.ToArray()
}

if ($Json) {
    $summary | ConvertTo-Json -Depth 6
} else {
    $checks | Format-Table status,required,id,name,detail -AutoSize | Out-String | Write-Host
    if ($failedRequired.Count -eq 0) {
        Write-Host "PREFLIGHT: PASS"
    } else {
        Write-Host "PREFLIGHT: FAIL - missing required dependencies: $($summary.missingRequired -join ', ')" -ForegroundColor Red
    }
}

if ($failedRequired.Count -gt 0) { exit 1 }
exit 0
