param(
    [string]$Profile = 'default',
    [string]$RuntimeSkillsPath = '',
    [switch]$LegacyCopyMode,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

if (-not $LegacyCopyMode) {
    Write-Host 'Codex uses junction registrations for this repo. Runtime edits already affect the source skills.'
    Write-Host 'Use -LegacyCopyMode only to import from an older copied Codex runtime.'
    return
}

function Get-RepoRoot {
    return (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}

function Get-DefaultRuntimeSkillsPath {
    if ($env:CODEX_HOME) {
        return (Join-Path $env:CODEX_HOME 'skills')
    }
    if ($HOME) {
        return (Join-Path $HOME '.codex\skills')
    }
    throw 'Cannot resolve Codex skills path. Set CODEX_HOME or pass -RuntimeSkillsPath.'
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

function Get-LinkTarget {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $null
    }
    $item = Get-Item -LiteralPath $Path -Force
    if (-not $item.LinkType) {
        return $null
    }
    if ($item.Target -is [array]) {
        return [string]$item.Target[0]
    }
    return [string]$item.Target
}

$repoRoot = Get-RepoRoot
$skillsRoot = Join-Path $repoRoot 'skills'
if (-not $RuntimeSkillsPath) {
    $RuntimeSkillsPath = Get-DefaultRuntimeSkillsPath
}
$RuntimeSkillsPath = [System.IO.Path]::GetFullPath($RuntimeSkillsPath)
$skills = Get-ProfileSkills -RepoRoot $repoRoot -ProfileName $Profile

if (-not (Test-Path -LiteralPath $RuntimeSkillsPath)) {
    throw "Codex runtime skills path not found: $RuntimeSkillsPath"
}

$backupRoot = $null
if (-not $NoBackup) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $repoRoot "generated\codex\source-backups\$stamp"
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
}

foreach ($skill in $skills) {
    if ($skill -eq '.system') {
        throw 'Refusing to sync .system skills.'
    }

    $runtimeSource = Join-Path $RuntimeSkillsPath $skill
    $repoTarget = Join-Path $skillsRoot $skill
    if (-not (Test-Path -LiteralPath (Join-Path $runtimeSource 'SKILL.md'))) {
        throw "Runtime skill missing SKILL.md: $runtimeSource"
    }
    if (Get-LinkTarget -Path $runtimeSource) {
        Write-Host "Skipping linked runtime skill: $skill"
        continue
    }

    if ((Test-Path -LiteralPath $repoTarget) -and -not $NoBackup) {
        Copy-Item -LiteralPath $repoTarget -Destination $backupRoot -Recurse -Force
    }

    if (Test-Path -LiteralPath $repoTarget) {
        Remove-Item -LiteralPath $repoTarget -Recurse -Force
    }
    Copy-Item -LiteralPath $runtimeSource -Destination $skillsRoot -Recurse -Force
    Write-Host "Synced legacy Codex runtime skill into repo: $skill"
}

Write-Host "Legacy Codex sync complete. Source repo: $skillsRoot"
