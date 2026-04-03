param(
    [switch]$InstallMissing = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$wingetArgs = @(
    '--accept-source-agreements',
    '--accept-package-agreements',
    '--silent'
)

function Add-ToSessionPath {
    param([string]$PathEntry)

    if ([string]::IsNullOrWhiteSpace($PathEntry) -or -not (Test-Path -LiteralPath $PathEntry)) {
        return
    }

    $parts = @($env:PATH -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts -notcontains $PathEntry) {
        $env:PATH = $PathEntry + ';' + $env:PATH
    }
}

function Resolve-Executable {
    param(
        [string[]]$CommandNames,
        [string[]]$CandidatePaths = @()
    )

    foreach ($candidate in $CandidatePaths) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    foreach ($name in $CommandNames) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
            return $command.Source
        }
    }

    return $null
}

function Ensure-WingetPackage {
    param(
        [string]$Id,
        [string]$DisplayName,
        [string[]]$CommandNames,
        [string[]]$CandidatePaths = @()
    )

    $existing = Resolve-Executable -CommandNames $CommandNames -CandidatePaths $CandidatePaths
    if ($null -ne $existing) {
        Add-ToSessionPath -PathEntry (Split-Path -Parent $existing)
        return [pscustomobject]@{
            installed = $true
            source = 'existing'
            path = $existing
        }
    }

    if (-not $InstallMissing) {
        return [pscustomobject]@{
            installed = $false
            source = 'missing'
            path = $null
        }
    }

    Write-Host "Installing $DisplayName via winget..."
    & winget install --id $Id -e @wingetArgs

    $resolved = Resolve-Executable -CommandNames $CommandNames -CandidatePaths $CandidatePaths
    if ($null -eq $resolved) {
        throw "Installed $DisplayName, but the executable could not be located."
    }

    Add-ToSessionPath -PathEntry (Split-Path -Parent $resolved)
    return [pscustomobject]@{
        installed = $true
        source = 'winget'
        path = $resolved
    }
}

function Ensure-PythonModule {
    param(
        [string]$PythonPath,
        [string]$ModuleName,
        [string]$PackageName = $ModuleName
    )

    & $PythonPath -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ModuleName') else 1)" 1>$null 2>$null
    if ($LASTEXITCODE -eq 0) {
        return [pscustomobject]@{
            installed = $true
            source = 'existing'
            module = $ModuleName
            package = $PackageName
        }
    }

    if (-not $InstallMissing) {
        return [pscustomobject]@{
            installed = $false
            source = 'missing'
            module = $ModuleName
            package = $PackageName
        }
    }

    Write-Host "Installing Python package $PackageName..."
    & $PythonPath -m pip install $PackageName | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install Python package $PackageName."
    }

    & $PythonPath -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('$ModuleName') else 1)" 1>$null 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "Installed $PackageName, but Python still cannot import module $ModuleName."
    }

    return [pscustomobject]@{
        installed = $true
        source = 'pip'
        module = $ModuleName
        package = $PackageName
    }
}

function Get-UninstallMatch {
    param([string]$DisplayNamePattern)

    $roots = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($root in $roots) {
        $match = Get-ItemProperty -Path $root -ErrorAction SilentlyContinue |
            Where-Object { $_.PSObject.Properties.Name -contains 'DisplayName' -and $_.DisplayName -like $DisplayNamePattern } |
            Select-Object -First 1
        if ($null -ne $match) {
            return $match
        }
    }

    return $null
}

function Get-WordInstallPath {
    $appPathKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    )

    foreach ($key in $appPathKeys) {
        $item = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if ($null -ne $item -and (Test-Path -LiteralPath $item.'(default)')) {
            return $item.'(default)'
        }
    }

    $candidates = @(
        'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
        'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Get-MathTypeInfo {
    $exeCandidates = @(
        'C:\Program Files (x86)\MathType\MathType.exe',
        'C:\Program Files\MathType\MathType.exe',
        'C:\Program Files\WIRIS\MathType\MathType.exe'
    )

    $exePath = $null
    foreach ($candidate in $exeCandidates) {
        if (Test-Path -LiteralPath $candidate) {
            $exePath = $candidate
            break
        }
    }

    $addinCandidates = @(
        (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\MathType Commands 2016.dotm'),
        (Join-Path $env:APPDATA 'Microsoft\Templates\MathType Commands 2016.dotm')
    )

    $addinPath = $null
    foreach ($candidate in $addinCandidates) {
        if (Test-Path -LiteralPath $candidate) {
            $addinPath = $candidate
            break
        }
    }

    $uninstallMatch = Get-UninstallMatch -DisplayNamePattern '*MathType*'

    return [pscustomobject]@{
        installed = ($null -ne $exePath -or $null -ne $addinPath -or $null -ne $uninstallMatch)
        exePath = $exePath
        addinPath = $addinPath
        uninstallName = if ($null -ne $uninstallMatch) { $uninstallMatch.DisplayName } else { $null }
    }
}

$python = Ensure-WingetPackage -Id 'Python.Python.3.12' -DisplayName 'Python 3.12' -CommandNames @('python', 'python.exe') -CandidatePaths @(
    'C:\Users\Admin\AppData\Local\Programs\Python\Python312\python.exe',
    'C:\Program Files\Python312\python.exe'
)
$pythonYaml = Ensure-PythonModule -PythonPath $python.path -ModuleName 'yaml' -PackageName 'PyYAML'

$git = Ensure-WingetPackage -Id 'Git.Git' -DisplayName 'Git' -CommandNames @('git', 'git.exe') -CandidatePaths @(
    'C:\Program Files\Git\cmd\git.exe',
    'C:\Program Files\Git\bin\git.exe'
)

$githubCli = Ensure-WingetPackage -Id 'GitHub.cli' -DisplayName 'GitHub CLI' -CommandNames @('gh', 'gh.exe') -CandidatePaths @(
    'C:\Program Files\GitHub CLI\gh.exe',
    'C:\Users\Admin\AppData\Local\Programs\GitHub CLI\gh.exe'
)

$wordPath = Get-WordInstallPath
$mathType = Get-MathTypeInfo

if ($null -eq $wordPath) {
    throw 'Microsoft Word was not found. Install Word before using this skill.'
}

if (-not $mathType.installed) {
    Write-Warning 'MathType was not found. Conversion can still run with native Word equations, but MathType-first conversion will be unavailable.'
}

[pscustomobject]@{
    python = $python
    pythonModules = [pscustomobject]@{
        yaml = $pythonYaml
    }
    git = $git
    githubCli = $githubCli
    word = [pscustomobject]@{
        installed = $true
        path = $wordPath
    }
    mathType = $mathType
}
