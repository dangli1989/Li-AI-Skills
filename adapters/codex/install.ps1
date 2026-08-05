param(
    [string]$Profile = 'default',
    [string]$RuntimeSkillsPath = '',
    [switch]$Force,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'

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

function Test-SamePath {
    param([string]$Left, [string]$Right)
    if (-not $Left -or -not $Right) {
        return $false
    }
    $leftFull = [System.IO.Path]::GetFullPath($Left).TrimEnd('\', '/')
    $rightFull = [System.IO.Path]::GetFullPath($Right).TrimEnd('\', '/')
    return [string]::Equals($leftFull, $rightFull, [System.StringComparison]::OrdinalIgnoreCase)
}

function Backup-ExistingEntry {
    param([string]$Path, [string]$BackupRoot)
    if (-not $BackupRoot) {
        return
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }
    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
    Copy-Item -LiteralPath $Path -Destination $BackupRoot -Recurse -Force
}

function Remove-ExistingEntry {
    param([string]$Path)
    $item = Get-Item -LiteralPath $Path -Force
    if ($item.LinkType) {
        Remove-Item -LiteralPath $Path -Force
    } else {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
}

function New-SkillRegistration {
    param([string]$Source, [string]$Target)
    try {
        New-Item -ItemType Junction -Path $Target -Target $Source | Out-Null
    } catch {
        New-Item -ItemType SymbolicLink -Path $Target -Target $Source | Out-Null
    }
}

$repoRoot = Get-RepoRoot
$skillsRoot = Join-Path $repoRoot 'skills'
if (-not $RuntimeSkillsPath) {
    $RuntimeSkillsPath = Get-DefaultRuntimeSkillsPath
}
$RuntimeSkillsPath = [System.IO.Path]::GetFullPath($RuntimeSkillsPath)
$skills = Get-ProfileSkills -RepoRoot $repoRoot -ProfileName $Profile

New-Item -ItemType Directory -Path $RuntimeSkillsPath -Force | Out-Null

$backupRoot = $null
if ($Force -and -not $NoBackup) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $repoRoot "generated\codex\backups\$stamp"
}

foreach ($skill in $skills) {
    if ($skill -eq '.system') {
        throw 'Refusing to register .system skills.'
    }

    $source = Join-Path $skillsRoot $skill
    $target = Join-Path $RuntimeSkillsPath $skill
    if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md'))) {
        throw "Source skill missing SKILL.md: $source"
    }

    if (Test-Path -LiteralPath $target) {
        $existingTarget = Get-LinkTarget -Path $target
        if (Test-SamePath $existingTarget $source) {
            Write-Host "Codex skill already registered: $skill"
            continue
        }

        if (-not $Force) {
            if ($existingTarget) {
                throw "Runtime skill '$skill' points elsewhere. Re-run with -Force to replace it: $existingTarget"
            }
            throw "Runtime skill '$skill' is a real directory or file. Re-run with -Force to back it up and replace it."
        }

        Backup-ExistingEntry -Path $target -BackupRoot $backupRoot
        Remove-ExistingEntry -Path $target
    }

    New-SkillRegistration -Source $source -Target $target
    Write-Host "Registered Codex skill: $skill"
}

Write-Host "Codex registration complete. Runtime: $RuntimeSkillsPath"
Write-Host 'Restart Codex to load changed skill registrations.'
