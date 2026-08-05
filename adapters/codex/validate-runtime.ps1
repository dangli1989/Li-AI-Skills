param(
    [string]$Profile = 'default',
    [string]$RuntimeSkillsPath = ''
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

$repoRoot = Get-RepoRoot
$skillsRoot = Join-Path $repoRoot 'skills'
if (-not $RuntimeSkillsPath) {
    $RuntimeSkillsPath = Get-DefaultRuntimeSkillsPath
}
$RuntimeSkillsPath = [System.IO.Path]::GetFullPath($RuntimeSkillsPath)
$skills = Get-ProfileSkills -RepoRoot $repoRoot -ProfileName $Profile
$errors = New-Object System.Collections.Generic.List[string]

foreach ($skill in $skills) {
    $source = Join-Path $skillsRoot $skill
    $runtimePath = Join-Path $RuntimeSkillsPath $skill
    if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md'))) {
        $errors.Add("Missing source skill: $skill")
        continue
    }
    if (-not (Test-Path -LiteralPath (Join-Path $runtimePath 'SKILL.md'))) {
        $errors.Add("Missing runtime skill: $skill")
        continue
    }

    $linkTarget = Get-LinkTarget -Path $runtimePath
    if (-not $linkTarget) {
        $errors.Add("Runtime skill is not a junction or symbolic link: $skill")
        continue
    }
    if (-not (Test-SamePath $linkTarget $source)) {
        $errors.Add("Runtime skill points to the wrong source: $skill -> $linkTarget")
    }
}

if (Test-Path -LiteralPath (Join-Path $RuntimeSkillsPath '.system')) {
    Write-Host 'Codex .system skills exist in runtime and are intentionally ignored.'
}

if ($errors.Count -gt 0) {
    $errors | ForEach-Object { Write-Host " - $_" -ForegroundColor Red }
    exit 1
}

Write-Host "Codex runtime validation passed for profile '$Profile'."
