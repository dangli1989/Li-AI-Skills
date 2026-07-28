param(
    [Parameter(Mandatory = $true)]
    [string]$DeckPath,

    [Parameter(Mandatory = $true)]
    [int]$SlideIndex,

    [Parameter(Mandatory = $true)]
    [string]$OutPath,

    [int]$Width = 1600,
    [int]$Height = 900
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DeckPath)) {
    throw "Deck not found: $DeckPath"
}

$chromeCandidates = @(
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
    'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
)

$browser = $chromeCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $browser) {
    throw "Chrome or Edge was not found."
}

$deck = Resolve-Path -LiteralPath $DeckPath
$deckRoot = Split-Path -Parent $deck
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("deck-slide-{0}.html" -f ([guid]::NewGuid()))
$profile = Join-Path ([System.IO.Path]::GetTempPath()) ("deck-render-profile-{0}" -f ([guid]::NewGuid()))
$outDir = Split-Path -Parent $OutPath
if (-not $outDir) {
    $outDir = (Get-Location).Path
}
$capturePath = Join-Path $outDir ("deck-render-{0}.png" -f ([guid]::NewGuid()))
New-Item -ItemType Directory -Path $profile -Force | Out-Null

$html = Get-Content -LiteralPath $deck -Raw
$baseHref = "file:///" + (($deckRoot -replace '\\', '/') + "/")
$html = $html -replace '<head>', "<head>`n  <base href=""$baseHref"">"

$slideNumber = -1
$html = [regex]::Replace($html, '<section\s+class="slide(?<classes>[^"]*)">', {
    param($match)
    $script:slideNumber += 1
    $extraClasses = $match.Groups['classes'].Value
    $extraClasses = $extraClasses -replace '\s*\bactive\b', ''
    $extraClasses = $extraClasses -replace '\s*\bprev\b', ''
    $extraClasses = $extraClasses.Trim()

    $classes = 'slide'
    if ($script:slideNumber -eq $SlideIndex) {
        $classes += ' active'
    }
    elseif ($script:slideNumber -lt $SlideIndex) {
        $classes += ' prev'
    }
    if ($extraClasses) {
        $classes += " $extraClasses"
    }

    return "<section class=""$classes"">"
})

# Static screenshots should not depend on deck JavaScript timing or animations.
$html = $html -replace '<script\s+src="jelly\.js"></script>', ''
$html = [regex]::Replace($html, '<script>\s*const slides = Array\.from\(document\.querySelectorAll\("\.slide"\)\);[\s\S]*?</script>', '')
$renderCss = @'
<style>
  .slide {
    opacity: 0 !important;
    visibility: hidden !important;
    transform: none !important;
    transition: none !important;
    animation: none !important;
    pointer-events: none !important;
  }
  .slide.active {
    opacity: 1 !important;
    visibility: visible !important;
    transform: none !important;
    pointer-events: auto !important;
  }
  .reveal,
  .jelly,
  .active .reveal {
    opacity: 1 !important;
    transform: none !important;
    transition: none !important;
    animation: none !important;
  }
</style>
'@
$html = $html -replace '</head>', "$renderCss</head>"
Set-Content -LiteralPath $tmp -Value $html -Encoding UTF8

try {
    $uri = "file:///" + ($tmp -replace '\\', '/')
    & $browser --headless=new --disable-gpu --disable-crash-reporter --disable-breakpad --disable-features=Crashpad --no-first-run --no-default-browser-check --user-data-dir="$profile" --hide-scrollbars "--window-size=$Width,$Height" "--screenshot=$capturePath" $uri | Out-Null
    if (-not (Test-Path -LiteralPath $capturePath)) {
        throw "Browser did not create screenshot: $capturePath"
    }
    Remove-Item -LiteralPath $OutPath -Force -ErrorAction SilentlyContinue
    Move-Item -LiteralPath $capturePath -Destination $OutPath -Force
    Get-Item -LiteralPath $OutPath
}
finally {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $capturePath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $profile -Recurse -Force -ErrorAction SilentlyContinue
}
