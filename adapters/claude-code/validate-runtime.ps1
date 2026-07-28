param(
    [string]$Profile = 'default',
    [string]$RuntimeSkillsPath = (Join-Path $env:USERPROFILE '.claude\skills')
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
$skills = Get-ProfileSkills -RepoRoot $repoRoot -ProfileName $Profile
$errors = New-Object System.Collections.Generic.List[string]

foreach ($skill in $skills) {
    $path = Join-Path $RuntimeSkillsPath $skill
    if (-not (Test-Path -LiteralPath (Join-Path $path 'SKILL.md'))) {
        $errors.Add("Missing runtime skill: $skill")
    }
}

if ($errors.Count -gt 0) {
    $errors | ForEach-Object { Write-Host " - $_" -ForegroundColor Red }
    exit 1
}

Write-Host "Claude Code runtime validation passed for profile '$Profile'."
