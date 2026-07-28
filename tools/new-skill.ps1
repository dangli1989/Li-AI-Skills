param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-z0-9][a-z0-9-]*$')]
    [string]$Id,

    [Parameter(Mandatory = $true)]
    [string]$Name,

    [Parameter(Mandatory = $true)]
    [string]$Description
)

$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$skillRoot = Join-Path $repoRoot "skills\$Id"

if (Test-Path -LiteralPath $skillRoot) {
    throw "Skill already exists: $Id"
}

New-Item -ItemType Directory -Path $skillRoot | Out-Null
New-Item -ItemType Directory -Path (Join-Path $skillRoot 'references') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $skillRoot 'scripts') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $skillRoot 'assets') | Out-Null

$skillMd = @"
---
name: $Id
description: $Description
---

# $Name

## Goal

Describe what this skill helps the agent do.

## Workflow

1. Add the concrete steps here.
"@

$metadata = @"
id: $Id
name: $Name
description: $Description
status: draft
owner: Li
entrypoint: SKILL.md
tags: []
resources:
  references: true
  scripts: true
  assets: true
supported_adapters:
  - codex
  - claude-code
"@

Set-Content -LiteralPath (Join-Path $skillRoot 'SKILL.md') -Value $skillMd -Encoding UTF8
Set-Content -LiteralPath (Join-Path $skillRoot 'skill.yaml') -Value $metadata -Encoding UTF8

Write-Host "Created new skill scaffold: $skillRoot"
