$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$skillsRoot = Join-Path $repoRoot 'skills'
$requiredDirs = @('skills', 'adapters', 'profiles', 'tools', 'generated', 'docs', 'tests')
$errors = New-Object System.Collections.Generic.List[string]

foreach ($dir in $requiredDirs) {
    $path = Join-Path $repoRoot $dir
    if (-not (Test-Path -LiteralPath $path)) {
        $errors.Add("Missing required directory: $dir")
    }
}

if (Test-Path -LiteralPath (Join-Path $skillsRoot '.system')) {
    $errors.Add('Do not store Codex .system skills in this repo.')
}

$forbiddenPathPatterns = @(
    '\\.matlab\\agentic-toolkits',
    '\\.codex\\skills\\.system'
)

$machineSpecificPathPatterns = @(
    ('C:' + '\\Users\\[^\\\s]+'),
    ('C:' + '/Users/' + '[^/\s]+'),
    ('C:' + '\\DevGit\b'),
    ('C:' + '/DevGit' + '\b'),
    ('\b' + 'li' + 'dang' + '\b')
)

$skillDirs = Get-ChildItem -Directory -LiteralPath $skillsRoot
foreach ($skill in $skillDirs) {
    $skillMd = Join-Path $skill.FullName 'SKILL.md'
    $skillYaml = Join-Path $skill.FullName 'skill.yaml'

    if (-not (Test-Path -LiteralPath $skillMd)) {
        $errors.Add("Missing SKILL.md: $($skill.Name)")
        continue
    }
    if (-not (Test-Path -LiteralPath $skillYaml)) {
        $errors.Add("Missing skill.yaml: $($skill.Name)")
    }

    $content = Get-Content -Raw -LiteralPath $skillMd
    if ($content -notmatch '(?s)^---\s*.*?name:\s*.+?.*?description:\s*.+?.*?---') {
        $errors.Add("SKILL.md frontmatter missing name/description: $($skill.Name)")
    }
    if ($content -match 'TODO|FIXME|\bTBD\b') {
        $errors.Add("Unresolved placeholder in SKILL.md: $($skill.Name)")
    }

    $allTextFiles = Get-ChildItem -Recurse -File -LiteralPath $skill.FullName |
        Where-Object { $_.Extension -in @('.md', '.yaml', '.yml', '.ps1', '.html', '.css', '.js', '.txt') }

    foreach ($file in $allTextFiles) {
        $text = Get-Content -Raw -LiteralPath $file.FullName
        foreach ($pattern in $forbiddenPathPatterns) {
            if ($text -match $pattern) {
                $errors.Add("Forbidden vendor/system path reference in $($file.FullName)")
            }
        }
    }
}

foreach ($adapter in @('codex', 'claude-code')) {
    $adapterRoot = Join-Path $repoRoot "adapters\$adapter"
    foreach ($fileName in @('adapter.yaml', 'install.ps1', 'sync-from-runtime.ps1', 'validate-runtime.ps1')) {
        $path = Join-Path $adapterRoot $fileName
        if (-not (Test-Path -LiteralPath $path)) {
            $errors.Add("Missing adapter file: adapters/$adapter/$fileName")
        }
    }
}

$repoTextFiles = Get-ChildItem -Recurse -File -LiteralPath $repoRoot |
    Where-Object {
        $_.FullName -notmatch '\\.git\\|\\generated\\' -and
        $_.Extension -in @('.md', '.yaml', '.yml', '.ps1', '.html', '.css', '.js', '.txt', '.gitignore')
    }

foreach ($file in $repoTextFiles) {
    $text = Get-Content -Raw -LiteralPath $file.FullName
    foreach ($pattern in $machineSpecificPathPatterns) {
        if ($text -match $pattern) {
            $relativePath = Resolve-Path -LiteralPath $file.FullName -Relative
            $errors.Add("Machine-specific path or username in $relativePath")
            break
        }
    }
}

if ($errors.Count -gt 0) {
    Write-Host 'Validation failed:' -ForegroundColor Red
    $errors | ForEach-Object { Write-Host " - $_" -ForegroundColor Red }
    exit 1
}

Write-Host "Validation passed: $($skillDirs.Count) skills, 2 adapters." -ForegroundColor Green
