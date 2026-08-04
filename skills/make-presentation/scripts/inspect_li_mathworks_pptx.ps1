param(
    [Parameter(Mandatory = $true)]
    [string]$DeckPath,

    [switch]$AsJson
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DeckPath)) {
    throw "Deck not found: $DeckPath"
}

$resolvedDeck = (Resolve-Path -LiteralPath $DeckPath).Path
$powerPoint = $null
$presentation = $null

function Get-PlaceholderInfo($shape) {
    try {
        return [pscustomobject]@{
            Name = $shape.Name
            Type = [int]$shape.PlaceholderFormat.Type
            Left = [double]$shape.Left
            Top = [double]$shape.Top
            Width = [double]$shape.Width
            Height = [double]$shape.Height
        }
    } catch {
        return $null
    }
}

try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($resolvedDeck, 0, 0, -1)

    $layouts = @()
    foreach ($layout in $presentation.SlideMaster.CustomLayouts) {
        $placeholders = @()
        foreach ($shape in $layout.Shapes) {
            $info = Get-PlaceholderInfo $shape
            if ($info -ne $null) {
                $placeholders += $info
            } elseif ($shape.Name -match 'Title|Content|Subtitle|Logo|Confidential|Copyright|Line') {
                $placeholders += [pscustomobject]@{
                    Name = $shape.Name
                    Type = $null
                    Left = [double]$shape.Left
                    Top = [double]$shape.Top
                    Width = [double]$shape.Width
                    Height = [double]$shape.Height
                }
            }
        }
        $layouts += [pscustomobject]@{
            Index = [int]$layout.Index
            Name = $layout.Name
            Placeholders = $placeholders
        }
    }

    $slides = @()
    foreach ($slide in $presentation.Slides) {
        $texts = @()
        $smallTextCount = 0
        $rectCount = 0
        foreach ($shape in $slide.Shapes) {
            try {
                if ($shape.AutoShapeType -eq 1) { $rectCount++ }
            } catch {}
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $text = $shape.TextFrame.TextRange.Text.Trim()
                    if ($text) { $texts += $text }
                    if ([double]$shape.TextFrame.TextRange.Font.Size -lt 12) { $smallTextCount++ }
                }
            } catch {}
        }
        $title = ''
        try { $title = $slide.Shapes.Title.TextFrame.TextRange.Text.Trim() } catch {}
        $slides += [pscustomobject]@{
            Index = [int]$slide.SlideIndex
            Layout = $slide.CustomLayout.Name
            Title = $title
            ShapeCount = [int]$slide.Shapes.Count
            TimingEffectCount = [int]$slide.TimeLine.MainSequence.Count
            RectShapeCount = $rectCount
            SmallTextShapeCount = $smallTextCount
            TextPreview = (($texts | Select-Object -First 4) -join ' | ')
        }
    }

    $result = [pscustomobject]@{
        DeckPath = $resolvedDeck
        SlideCount = [int]$presentation.Slides.Count
        Layouts = $layouts
        Slides = $slides
    }

    if ($AsJson) {
        $result | ConvertTo-Json -Depth 8
    } else {
        $result | Format-List
    }
}
finally {
    if ($presentation -ne $null) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($powerPoint -ne $null) {
        $powerPoint.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    }
}
