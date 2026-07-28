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

New-Item -ItemType Directory -Path $RuntimeSkillsPath -Force | Out-Null

if (-not $NoBackup) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $repoRoot "generated\codex\backups\$stamp"
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
}

foreach ($skill in $skills) {
    if ($skill -eq '.system') {
        throw 'Refusing to install .system skills.'
    }

    $source = Join-Path $skillsRoot $skill
    $target = Join-Path $RuntimeSkillsPath $skill
    if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md'))) {
        throw "Source skill missing SKILL.md: $source"
    }

    if ((Test-Path -LiteralPath $target) -and -not $NoBackup) {
        Copy-Item -LiteralPath $target -Destination $backupRoot -Recurse -Force
    }

    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force
    }
    Copy-Item -LiteralPath $source -Destination $RuntimeSkillsPath -Recurse -Force
    Write-Host "Installed Codex skill: $skill"
}

Write-Host "Codex install complete. Runtime: $RuntimeSkillsPath"
