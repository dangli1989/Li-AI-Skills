param(
    [Parameter(Mandatory = $true)]
    [string]$DeckPath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DeckPath)) {
    throw "Deck not found: $DeckPath"
}

$resolved = Resolve-Path -LiteralPath $DeckPath
$root = Split-Path -Parent $resolved
$html = Get-Content -LiteralPath $resolved -Raw

$slideCount = ([regex]::Matches($html, '<section class="slide')).Count
$timerMatches = [regex]::Matches($html, 'Slide [0-9]+ / ([0-9]+)')
$timerCount = $timerMatches.Count
$declaredTotal = if ($timerMatches.Count -gt 0) { [int]$timerMatches[0].Groups[1].Value } else { 0 }

$srcs = [regex]::Matches($html, '(?:src|href)="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
$missing = @()
foreach ($src in $srcs) {
    if ($src -match '^(https?:|mailto:|matlab:|#)') { continue }
    $path = Join-Path $root $src
    if (-not (Test-Path -LiteralPath $path)) { $missing += $src }
}

[pscustomobject]@{
    Deck = $resolved.Path
    SlideCount = $slideCount
    TimerCount = $timerCount
    DeclaredTotal = $declaredTotal
    CountersMatch = ($slideCount -eq $timerCount -and $slideCount -eq $declaredTotal)
    MissingAssetCount = $missing.Count
    MissingAssets = ($missing -join '; ')
}
