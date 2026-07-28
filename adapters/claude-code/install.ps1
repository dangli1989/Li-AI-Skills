param(
    [string]$Profile = 'default',
    [string]$RuntimeSkillsPath = (Join-Path $env:USERPROFILE '.claude\skills'),
    [switch]$ExportOnly
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
$bundleRoot = Join-Path $repoRoot 'generated\claude-code\skills'
$skills = Get-ProfileSkills -RepoRoot $repoRoot -ProfileName $Profile

if (Test-Path -LiteralPath $bundleRoot) {
    Remove-Item -LiteralPath $bundleRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $bundleRoot -Force | Out-Null

foreach ($skill in $skills) {
    $source = Join-Path $skillsRoot $skill
    if (-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md'))) {
        throw "Source skill missing SKILL.md: $source"
    }
    Copy-Item -LiteralPath $source -Destination $bundleRoot -Recurse -Force
    Write-Host "Exported Claude Code skill bundle: $skill"
}

if (-not $ExportOnly) {
    New-Item -ItemType Directory -Path $RuntimeSkillsPath -Force | Out-Null
    foreach ($skill in $skills) {
        $source = Join-Path $bundleRoot $skill
        $target = Join-Path $RuntimeSkillsPath $skill
        if (Test-Path -LiteralPath $target) {
            Remove-Item -LiteralPath $target -Recurse -Force
        }
        Copy-Item -LiteralPath $source -Destination $RuntimeSkillsPath -Recurse -Force
        Write-Host "Installed Claude Code skill: $skill"
    }
    Write-Host "Claude Code install complete. Runtime: $RuntimeSkillsPath"
} else {
    Write-Host "Claude Code export complete. Bundle: $bundleRoot"
}
