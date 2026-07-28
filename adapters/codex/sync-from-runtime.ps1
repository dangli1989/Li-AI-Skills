param(
    [string]$Profile = 'default',
    [string]$RuntimeSkillsPath = (Join-Path $env:USERPROFILE '.codex\skills'),
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

function Get-RepoRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}

function Get-ProfileSkills {
    param([string]$RepoRoot, [string]$ProfileName)
    $profilePath = Join-Path $RepoRoot "profiles\$ProfileName.yaml"
    if (-not (Test-Path -LiteralPath $profilePath)) {
        throw "Profile not found: $profilePath"
    }

    $skills = New-Object System.Collections.Generic.List[string]
    foreach ($line in Get-Content -LiteralPath $profilePath) {
        if ($line -match '^\s*-\s*([a-z0-9][a-z0-9-]*)\s*$') {
            $skills.Add($Matches[1])
        }
    }
    return $skills
}

$repoRoot = Get-RepoRoot
$skillsRoot = Join-Path $repoRoot 'skills'
$skills = Get-ProfileSkills -RepoRoot $repoRoot -ProfileName $Profile

if (-not (Test-Path -LiteralPath $RuntimeSkillsPath)) {
    throw "Codex runtime skills path not found: $RuntimeSkillsPath"
}

if (-not $NoBackup) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $repoRoot "generated\codex\source-backups\$stamp"
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
}

foreach ($skill in $skills) {
    if ($skill -eq '.system') {
        throw 'Refusing to sync .system skills.'
    }

    $source = Join-Path $RuntimeSkillsPath $skill
    $target = Join-Path $skillsRoot $skill
    if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md'))) {
        throw "Runtime skill missing SKILL.md: $source"
    }

    if ((Test-Path -LiteralPath $target) -and -not $NoBackup) {
        Copy-Item -LiteralPath $target -Destination $backupRoot -Recurse -Force
    }

    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
    Copy-Item -LiteralPath $source -Destination $skillsRoot -Recurse -Force
    Write-Host "Synced Codex runtime skill into repo: $skill"
}

Write-Host "Codex sync complete. Source repo: $skillsRoot"
