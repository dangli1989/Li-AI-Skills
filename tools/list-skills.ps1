$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$skillsRoot = Join-Path $repoRoot 'skills'

Get-ChildItem -Directory -LiteralPath $skillsRoot | ForEach-Object {
    $skillPath = $_.FullName
    $skillFile = Join-Path $skillPath 'SKILL.md'
    $metadataFile = Join-Path $skillPath 'skill.yaml'

    [pscustomobject]@{
        Id = $_.Name
        HasSkillMd = Test-Path -LiteralPath $skillFile
        HasMetadata = Test-Path -LiteralPath $metadataFile
        Path = $skillPath
    }
} | Format-Table -AutoSize
