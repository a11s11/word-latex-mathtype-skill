param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$RepoName = 'word-latex-mathtype-skill',
    [ValidateSet('public', 'private')]
    [string]$Visibility = 'public',
    [switch]$SkipSetup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

if (-not $SkipSetup) {
    & (Join-Path $PSScriptRoot 'setup_environment.ps1') | Out-Host
}

$gitPath = Resolve-Executable -CommandNames @('git', 'git.exe') -CandidatePaths @(
    'C:\Program Files\Git\cmd\git.exe',
    'C:\Program Files\Git\bin\git.exe'
)

$ghPath = Resolve-Executable -CommandNames @('gh', 'gh.exe') -CandidatePaths @(
    'C:\Program Files\GitHub CLI\gh.exe',
    'C:\Users\Admin\AppData\Local\Programs\GitHub CLI\gh.exe'
)

if ($null -eq $gitPath) {
    throw 'Git was not found. Run setup_environment.ps1 first.'
}

if ($null -eq $ghPath) {
    throw 'GitHub CLI was not found. Run setup_environment.ps1 first.'
}

$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
Push-Location $RepoRoot

try {
    if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot '.git'))) {
        & $gitPath init -b main | Out-Host
    }

    & $gitPath branch -M main | Out-Host

    & cmd /c "`"$ghPath`" auth status >nul 2>nul"
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'Launching browser-based GitHub login...'
        & $ghPath auth login --web --git-protocol https | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw 'GitHub authentication failed or was cancelled.'
        }
    }

    $owner = ((& $ghPath api user --jq .login) | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($owner)) {
        throw 'Could not resolve the authenticated GitHub username.'
    }

    $configuredName = ((& $gitPath config user.name 2>$null) | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($configuredName)) {
        $displayName = ((& $ghPath api user --jq '.name // .login') | Out-String).Trim()
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = $owner
        }
        & $gitPath config user.name $displayName | Out-Host
    }

    $configuredEmail = ((& $gitPath config user.email 2>$null) | Out-String).Trim()
    if ([string]::IsNullOrWhiteSpace($configuredEmail)) {
        $email = ((& $ghPath api user --jq '.email // empty') | Out-String).Trim()
        if ([string]::IsNullOrWhiteSpace($email)) {
            $email = "$owner@users.noreply.github.com"
        }
        & $gitPath config user.email $email | Out-Host
    }

    & $gitPath add . | Out-Host
    $status = (& $gitPath status --porcelain)
    if (-not [string]::IsNullOrWhiteSpace(($status | Out-String))) {
        & $gitPath commit -m 'Initial skill release' | Out-Host
    }

    & cmd /c "`"$ghPath`" repo view $owner/$RepoName >nul 2>nul"
    if ($LASTEXITCODE -ne 0) {
        & $ghPath repo create $RepoName "--$Visibility" --source $RepoRoot --remote origin | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create the GitHub repository '$owner/$RepoName'."
        }
    } else {
        $remotes = @(& $gitPath remote)
        if ($remotes -notcontains 'origin') {
            & $gitPath remote add origin "https://github.com/$owner/$RepoName.git" | Out-Host
        }
    }

    & $gitPath push -u origin main | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw 'git push failed.'
    }

    [pscustomobject]@{
        owner = $owner
        repository = $RepoName
        remote = "https://github.com/$owner/$RepoName"
    }
}
finally {
    Pop-Location
}
