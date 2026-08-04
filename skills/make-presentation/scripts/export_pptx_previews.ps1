param(
    [Parameter(Mandatory = $true)]
    [string]$DeckPath,

    [Parameter(Mandatory = $true)]
    [string]$OutDir,

    [string]$Slides = 'all',

    [int]$Width = 1600,

    [int]$Height = 900
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DeckPath)) {
    throw "Deck not found: $DeckPath"
}

$resolvedDeck = (Resolve-Path -LiteralPath $DeckPath).Path
$resolvedOutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
New-Item -ItemType Directory -Force -Path $resolvedOutDir | Out-Null

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($resolvedDeck, 0, 0, -1)

    if ($Slides -eq 'all') {
        $slideNumbers = 1..$presentation.Slides.Count
    } else {
        $slideNumbers = $Slides.Split(',') | ForEach-Object { [int]$_.Trim() }
    }

    foreach ($slideNumber in $slideNumbers) {
        if ($slideNumber -lt 1 -or $slideNumber -gt $presentation.Slides.Count) {
            throw "Slide $slideNumber is outside deck range 1..$($presentation.Slides.Count)"
        }
        $output = Join-Path $resolvedOutDir ("slide-{0:D2}.png" -f $slideNumber)
        $presentation.Slides.Item($slideNumber).Export($output, 'PNG', $Width, $Height)
        Write-Host "Exported preview: $output"
    }
}
finally {
    if ($presentation -ne $null) {
        try { $presentation.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
    }
    if ($powerPoint -ne $null) {
        try { $powerPoint.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null } catch {}
    }
}
