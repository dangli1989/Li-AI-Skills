param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [ValidateSet('public', 'confidential')]
    [string]$TemplateKind = 'public',

    [string]$GroundTruthPath,

    [string]$SpeechPath,

    [switch]$ExportPreview
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SpecPath)) {
    throw "Spec not found: $SpecPath"
}

$skillRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$templatePath = Join-Path $skillRoot "assets\li-mathworks-presentation\pptx\$TemplateKind.pptx"
if (-not (Test-Path -LiteralPath $templatePath)) {
    throw "Template not found: $templatePath"
}

$resolvedSpec = (Resolve-Path -LiteralPath $SpecPath).Path
$script:SpecDir = Split-Path -Parent $resolvedSpec
$script:ProgressLabels = @()
$spec = Get-Content -Raw -LiteralPath $resolvedSpec | ConvertFrom-Json
$resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$outputDir = Split-Path -Parent $resolvedOutput
if ($outputDir) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}

function RgbLong([int]$r, [int]$g, [int]$b) {
    return $r + ($g -shl 8) + ($b -shl 16)
}

$Palette = @{
    Blue = RgbLong 0 118 168
    DeepBlue = RgbLong 0 75 135
    Cyan = RgbLong 0 169 224
    Orange = RgbLong 215 136 36
    Yellow = RgbLong 242 169 0
    Red = RgbLong 183 48 44
    Green = RgbLong 72 162 63
    Ink = RgbLong 31 31 31
    Muted = RgbLong 99 101 105
    LightBlue = RgbLong 236 245 248
    LightCyan = RgbLong 232 247 251
    LightYellow = RgbLong 255 246 218
    LightGreen = RgbLong 237 247 235
    LightRed = RgbLong 250 234 233
    LightGray = RgbLong 242 244 245
}

$LightFills = @($Palette.LightBlue, $Palette.LightYellow, $Palette.LightGreen, $Palette.LightCyan, $Palette.LightGray, $Palette.LightRed)
$AccentLines = @($Palette.Blue, $Palette.Orange, $Palette.Green, $Palette.Cyan, $Palette.Muted, $Palette.Red)

$BannedVisibleTextPatterns = @(
    '(?i)use progressive reveal',
    '(?i)speaker note',
    '(?i)presenter note',
    '(?i)build note',
    '(?i)click to',
    '(?i)animate this',
    '(?i)visual check',
    '(?i)\b(the workshop|the close|the deck|the slide)\s+should\b'
)

function Test-VisibleText($node, [string]$path) {
    if ($null -eq $node) { return }
    if ($node -is [string]) {
        foreach ($pattern in $BannedVisibleTextPatterns) {
            if ($node -match $pattern) {
                throw "Visible slide text appears to contain build/meta language at $path`: '$node'"
            }
        }
        return
    }
    if ($node.GetType().IsValueType) { return }
    if ($node -is [System.Collections.IEnumerable] -and -not ($node -is [string])) {
        $index = 0
        foreach ($item in $node) {
            Test-VisibleText $item "${path}[$index]"
            $index++
        }
        return
    }
    foreach ($property in $node.PSObject.Properties) {
        if ($property.Name -in @('reason', 'talkTrack', 'builds', 'code', 'coverTalkTrack', 'sourceFacts', 'missingInputs', 'assetStatus', 'placeholderStatus', 'audience', 'purpose')) { continue }
        Test-VisibleText $property.Value "$path.$($property.Name)"
    }
}

function Get-ProgressLabels($specNode) {
    if ($specNode.progressLabels) {
        return @($specNode.progressLabels | ForEach-Object { [string]$_ })
    }
    return @()
}

function Get-LocatorItems($slideSpec) {
    if ($slideSpec.locator) {
        return @($slideSpec.locator)
    }
    if ($slideSpec.activeProgressLabel -and $script:ProgressLabels.Count -gt 0) {
        $activeLabel = [string]$slideSpec.activeProgressLabel
        return @($script:ProgressLabels | ForEach-Object {
            [pscustomobject]@{
                label = $_
                active = ($_ -eq $activeLabel)
            }
        })
    }
    return @()
}

function Resolve-AssetPath([string]$path) {
    if (-not $path) { return '' }
    if ([System.IO.Path]::IsPathRooted($path)) { return $path }
    return (Join-Path $script:SpecDir $path)
}

function Test-Spec($specNode) {
    $script:ProgressLabels = @(Get-ProgressLabels $specNode)
    $progressSet = @{}
    foreach ($label in $script:ProgressLabels) {
        $key = $label.ToLowerInvariant()
        if ($progressSet.ContainsKey($key)) {
            throw "Top-level progressLabels contains duplicate label '$label'. Use one source of truth for section/progress labels."
        }
        $progressSet[$key] = $true
    }

    $slideIndex = 1
    foreach ($slideSpec in @($specNode.slides)) {
        if (-not $slideSpec.reason) {
            throw "Slide spec $slideIndex is missing reason. Record why the slide exists before generating."
        }
        if (-not $slideSpec.talkTrack) {
            throw "Slide spec $slideIndex is missing talkTrack. Generate the speech/talk track before styling."
        }

        $locator = @(Get-LocatorItems $slideSpec)
        if ($locator.Count -gt 0) {
            $seen = @{}
            $activeCount = 0
            foreach ($item in $locator) {
                $label = if ($item -is [string]) { [string]$item } else { [string]$item.label }
                $key = $label.ToLowerInvariant()
                if ($seen.ContainsKey($key)) {
                    throw "Slide spec $slideIndex has duplicate progress/locator label '$label'. Use top-level progressLabels plus activeProgressLabel instead of hand-entering repeated labels."
                }
                $seen[$key] = $true
                try {
                    if ([bool]$item.active) { $activeCount++ }
                } catch {}
            }
            if ($activeCount -ne 1) {
                throw "Slide spec $slideIndex has $activeCount active progress/locator labels; expected exactly one."
            }
        }

        if ($slideSpec.activeProgressLabel -and $script:ProgressLabels.Count -gt 0) {
            $activeKey = ([string]$slideSpec.activeProgressLabel).ToLowerInvariant()
            if (-not $progressSet.ContainsKey($activeKey)) {
                throw "Slide spec $slideIndex activeProgressLabel '$($slideSpec.activeProgressLabel)' is not in top-level progressLabels."
            }
        }

        if ([string]$slideSpec.type -in @('image-evidence', 'screenshot-evidence', 'model-screenshot', 'concept-artifact')) {
            foreach ($image in @($slideSpec.images)) {
                $imagePath = if ($image -is [string]) { [string]$image } else { [string]$image.path }
                $resolvedImage = Resolve-AssetPath $imagePath
                $isPlaceholder = $false
                try { $isPlaceholder = [bool]$image.placeholder } catch {}
                if (-not $isPlaceholder -and -not (Test-Path -LiteralPath $resolvedImage)) {
                    throw "Slide spec $slideIndex image not found: $imagePath"
                }
            }
        }
        $slideIndex++
    }
}

function Get-SemanticColors([string]$role, [int]$index) {
    $r = $role.ToLowerInvariant()
    if ($r -match 'legacy|current') { return @{ Fill = $Palette.LightGray; Line = $Palette.Muted } }
    if ($r -match 'model|workflow|architecture|scale') { return @{ Fill = $Palette.LightBlue; Line = $Palette.Blue } }
    if ($r -match 'conversion|action|next|resolve') { return @{ Fill = $Palette.LightYellow; Line = $Palette.Orange } }
    if ($r -match 'verification|validated|pass') { return @{ Fill = $Palette.LightGreen; Line = $Palette.Green } }
    if ($r -match 'risk|block|blocked|blocker|gap') { return @{ Fill = $Palette.LightRed; Line = $Palette.Red } }
    if ($r -match 'review|highlight|pending') { return @{ Fill = $Palette.LightYellow; Line = $Palette.Yellow } }
    return @{ Fill = $LightFills[$index % $LightFills.Count]; Line = $AccentLines[$index % $AccentLines.Count] }
}

function Get-Layout($presentation, [string]$name) {
    foreach ($layout in $presentation.SlideMaster.CustomLayouts) {
        if ($layout.Name -eq $name) { return $layout }
    }
    throw "Template layout not found: $name"
}

function Find-Placeholder($slide, [string[]]$namePatterns, [int[]]$types) {
    foreach ($shape in $slide.Shapes) {
        foreach ($pattern in $namePatterns) {
            try {
                if ($shape.Name -match $pattern) { return $shape }
            } catch {}
        }
    }
    foreach ($shape in $slide.Shapes) {
        try {
            if ($types -contains [int]$shape.PlaceholderFormat.Type) { return $shape }
        } catch {}
    }
    $best = $null
    $bestArea = 0
    foreach ($shape in $slide.Shapes) {
        try {
            if (-not $shape.HasTextFrame) { continue }
            $isTitleRequest = ($namePatterns -contains 'Title')
            $isContentRequest = ($namePatterns -contains 'Content' -or $namePatterns -contains 'Left' -or $namePatterns -contains 'Right')
            if ($isTitleRequest -and $shape.Top -gt 140) { continue }
            if ($isContentRequest -and $shape.Top -lt 100) { continue }
            $area = [double]$shape.Width * [double]$shape.Height
            if ($area -gt $bestArea) {
                $best = $shape
                $bestArea = $area
            }
        } catch {}
    }
    if ($best -ne $null) { return $best }
    return $null
}

function Find-StrictPlaceholder($slide, [string[]]$namePatterns, [int[]]$types) {
    foreach ($shape in $slide.Shapes) {
        foreach ($pattern in $namePatterns) {
            try {
                if ($shape.Name -match $pattern) { return $shape }
            } catch {}
        }
    }
    foreach ($shape in $slide.Shapes) {
        try {
            if ($types -contains [int]$shape.PlaceholderFormat.Type) { return $shape }
        } catch {}
    }
    return $null
}

function Set-Text($shape, [string]$text, [double]$size = 0, [bool]$bold = $false, [int]$color = 0) {
    if ($shape -eq $null) { return }
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.WordWrap = -1
    if ($size -gt 0) { $shape.TextFrame.TextRange.Font.Size = $size }
    if ($bold) { $shape.TextFrame.TextRange.Font.Bold = -1 }
    if ($color -ne 0) { $shape.TextFrame.TextRange.Font.Color.RGB = $color }
}

function Get-ShapeText($shape) {
    try {
        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
            return ([string]$shape.TextFrame.TextRange.Text).Trim()
        }
    } catch {}
    return ''
}

function Remove-EmptyTextShapes($presentation) {
    foreach ($slide in $presentation.Slides) {
        for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
            $shape = $slide.Shapes.Item($i)
            $hasTextFrame = $false
            try { $hasTextFrame = [bool]$shape.HasTextFrame } catch {}
            if (-not $hasTextFrame) { continue }

            $text = Get-ShapeText $shape
            if ($text.Length -gt 0) { continue }

            $placeholderType = $null
            try { $placeholderType = [int]$shape.PlaceholderFormat.Type } catch {}

            $isTextBoxShape = $false
            try { $isTextBoxShape = ([int]$shape.Type -eq 17) } catch {}
            try {
                if ($isTextBoxShape -or $placeholderType -in @(2, 4, 7)) {
                    $shape.Delete()
                }
            } catch {}
        }
    }
}

function Set-CoverText($cover, $specNode) {
    $title = Find-StrictPlaceholder $cover @('Title') @(1,3)
    if ($title -eq $null) { throw 'Title Slide layout does not expose a title placeholder.' }
    Set-Text $title ([string]$specNode.title) 0 $false 0

    $subtitle = Find-StrictPlaceholder $cover @('Subtitle') @(4)
    $subtitleText = [string]$specNode.subtitle
    $authorLine = [string]$specNode.authorLine
    $dateLine = ''
    if ($specNode.dateLine) { $dateLine = [string]$specNode.dateLine }
    elseif ($specNode.date) { $dateLine = [string]$specNode.date }

    if ($subtitleText) {
        if ($subtitle -eq $null) { throw 'Title Slide layout does not expose a subtitle placeholder.' }
        Set-Text $subtitle $subtitleText 0 $false 0
        if ($authorLine) {
            $author = Add-TransparentTextBox $cover $authorLine 72 300 520 22 14.5 $false $Palette.Ink
            Set-LiPatternTag $author 'CoverAuthor'
        }
        if ($dateLine) {
            $date = Add-TransparentTextBox $cover $dateLine 72 323 520 18 11.5 $false $Palette.Muted
            Set-LiPatternTag $date 'CoverDate'
        }
    } elseif ($subtitle -ne $null) {
        $templateSubtitle = @($authorLine, $dateLine) | Where-Object { $_ }
        Set-Text $subtitle ($templateSubtitle -join "`r") 0 $false 0
    }
}

function Add-Notes($slide, $slideSpec) {
    $noteLines = New-Object System.Collections.Generic.List[string]
    if ($slideSpec.reason) {
        $noteLines.Add("Reason: $($slideSpec.reason)") | Out-Null
    }
    if ($slideSpec.talkTrack) {
        $noteLines.Add("Talk track: $($slideSpec.talkTrack)") | Out-Null
    }
    if ($slideSpec.builds) {
        $noteLines.Add("Builds:") | Out-Null
        foreach ($build in @($slideSpec.builds)) {
            $noteLines.Add("- $build") | Out-Null
        }
    }
    if ($noteLines.Count -eq 0) { return }
    try {
        $notesText = ($noteLines -join "`r")
        $notesPage = $slide.NotesPage
        foreach ($shape in $notesPage.Shapes) {
            if ($shape.PlaceholderFormat.Type -eq 2 -and $shape.HasTextFrame) {
                $shape.TextFrame.TextRange.Text = $notesText
                return
            }
        }
    } catch {}
}

function Add-TransparentTextBox($slide, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [double]$size, [bool]$bold, [int]$color, [switch]$AllowEmpty) {
    if (-not $AllowEmpty -and [string]::IsNullOrWhiteSpace($text)) { return $null }
    $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.MarginLeft = 0
    $shape.TextFrame.MarginRight = 0
    $shape.TextFrame.MarginTop = 0
    $shape.TextFrame.MarginBottom = 0
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.TextRange.Font.Name = 'Arial'
    $shape.TextFrame.TextRange.Font.Size = $size
    $shape.TextFrame.TextRange.Font.Bold = if ($bold) { -1 } else { 0 }
    $shape.TextFrame.TextRange.Font.Color.RGB = $color
    return $shape
}

function Set-BodyParagraphs($shape, $items, [double]$fontSize = 16) {
    if ($shape -eq $null) { return }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($item in $items) {
        if ($item -is [string]) {
            $lines.Add($item) | Out-Null
        } else {
            $label = [string]$item.label
            $detail = [string]$item.detail
            if ($detail) {
                $lines.Add("$label - $detail") | Out-Null
            } else {
                $lines.Add($label) | Out-Null
            }
        }
    }
    $shape.TextFrame.TextRange.Text = ($lines -join "`r")
    $shape.TextFrame.TextRange.Font.Size = $fontSize
    $shape.TextFrame.TextRange.Font.Color.RGB = $Palette.Ink
    $shape.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1
}

function Add-Fade($slide, $shape, [int]$trigger = 1) {
    try {
        $effect = $slide.TimeLine.MainSequence.AddEffect($shape, 10, 0, $trigger)
        $effect.Timing.Duration = 0.5
    } catch {}
}

function Add-Wipe($slide, $shape, [int]$trigger = 1, [int]$direction = 1) {
    try {
        $effect = $slide.TimeLine.MainSequence.AddEffect($shape, 22, 0, $trigger)
        $effect.Timing.Duration = 0.5
        try { $effect.EffectParameters.Direction = $direction } catch {}
    } catch {
        Add-Fade $slide $shape $trigger
    }
}

function Add-RoundedCard($slide, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [int]$index, [double]$fontSize = 17, [switch]$NoAnimation) {
    $shape = $slide.Shapes.AddShape(5, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = $LightFills[$index % $LightFills.Count]
    $shape.Fill.Transparency = 0.08
    $shape.Line.ForeColor.RGB = $AccentLines[$index % $AccentLines.Count]
    $shape.Line.Weight = 0.9
    $shape.TextFrame.MarginLeft = 12
    $shape.TextFrame.MarginRight = 12
    $shape.TextFrame.MarginTop = 8
    $shape.TextFrame.MarginBottom = 8
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.Font.Size = $fontSize
    $shape.TextFrame.TextRange.Font.Bold = -1
    $shape.TextFrame.TextRange.Font.Color.RGB = $Palette.Ink
    if (-not $NoAnimation) { Add-Fade $slide $shape 1 }
    return $shape
}

function Add-Arrow($slide, [double]$left, [double]$top, [double]$width, [double]$height, [switch]$NoAnimation) {
    $shape = $slide.Shapes.AddShape(33, $left, $top, $width, $height)
    $shape.Fill.ForeColor.RGB = $Palette.Orange
    $shape.Line.Visible = 0
    if (-not $NoAnimation) { Add-Wipe $slide $shape 1 1 }
    return $shape
}

function Group-Shapes($slide, $shapes, [string]$name) {
    $names = @()
    foreach ($shape in $shapes) {
        try { $names += $shape.Name } catch {}
    }
    if ($names.Count -lt 2) { return $null }
    try {
        $group = $slide.Shapes.Range($names).Group()
        $group.Name = $name
        return $group
    } catch {
        return $null
    }
}

function Set-LiPatternTag($shape, [string]$pattern) {
    if ($shape -eq $null -or -not $pattern) { return }
    try { $shape.AlternativeText = "LiPattern:$pattern" } catch {}
    try {
        if ($shape.Name -notmatch '^Li_') {
            $shape.Name = "Li_$pattern`_$($shape.Id)"
        }
    } catch {}
}

function Add-PatternGroup($slide, $shapes, [string]$pattern, [string]$name) {
    $group = Group-Shapes $slide $shapes $name
    if ($group -ne $null) {
        try { $group.Line.Visible = 0 } catch {}
        Set-LiPatternTag $group $pattern
        return $group
    }
    foreach ($shape in $shapes) { Set-LiPatternTag $shape $pattern }
    return $null
}

function Add-TextBoxInShape($shape, [string]$text, [double]$fontSize, [bool]$bold, [int]$color) {
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.WordWrap = -1
    $shape.TextFrame.MarginLeft = 10
    $shape.TextFrame.MarginRight = 10
    $shape.TextFrame.MarginTop = 6
    $shape.TextFrame.MarginBottom = 6
    $shape.TextFrame.TextRange.Font.Name = 'Arial'
    $shape.TextFrame.TextRange.Font.Size = $fontSize
    $shape.TextFrame.TextRange.Font.Bold = if ($bold) { -1 } else { 0 }
    $shape.TextFrame.TextRange.Font.Color.RGB = $color
}

function Add-ProgressRail($slide, $items, [string]$activeLabel, [double]$left, [double]$top, [double]$width, [double]$rowHeight) {
    $shapes = @()
    $index = 0
    foreach ($item in @($items)) {
        $label = if ($item -is [string]) { [string]$item } else { [string]$item.label }
        $active = $false
        try { $active = [bool]$item.active } catch {}
        if ($activeLabel) { $active = ($label -eq $activeLabel) }
        $colors = if ($active) {
            @{ Fill = $Palette.LightGreen; Line = $Palette.Green; Text = $Palette.Ink }
        } else {
            @{ Fill = $Palette.LightGray; Line = $Palette.Orange; Text = $Palette.Muted }
        }
        $shape = $slide.Shapes.AddShape(5, $left, ($top + $index * ($rowHeight + 4)), $width, $rowHeight)
        $shape.Fill.ForeColor.RGB = $colors.Fill
        $shape.Fill.Transparency = if ($active) { 0.02 } else { 0.18 }
        $shape.Line.ForeColor.RGB = $colors.Line
        $shape.Line.Weight = if ($active) { 1.5 } else { 0.8 }
        Add-TextBoxInShape $shape $label 10.5 $active $colors.Text
        $shapes += $shape
        $index++
    }
    $group = Add-PatternGroup $slide $shapes 'ProgressRail' 'Li_ProgressRail'
    if ($group -ne $null) { Add-Fade $slide $group 1 }
    return $group
}

function Get-ContentPlaceholders($slide) {
    $items = @()
    foreach ($shape in $slide.Shapes) {
        try {
            $placeholderType = [int]$shape.PlaceholderFormat.Type
            if ($placeholderType -eq 7 -or $placeholderType -eq 2) {
                $items += $shape
            }
        } catch {}
    }
    return @($items | Sort-Object { [double]$_.Left })
}

function Add-AgendaSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    $title = Find-Placeholder $slide @('Title') @(1,3)
    if ($title -eq $null) {
        $title = Add-TransparentTextBox $slide ([string]$slideSpec.title) 48 37 848 90 28 $true $Palette.DeepBlue
    } else {
        Set-Text $title ([string]$slideSpec.title) 0 $false 0
    }
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -eq $null) {
        $content = Add-TransparentTextBox $slide '' 48 126 848 366 18 $false $Palette.Ink -AllowEmpty
    }
    Set-BodyParagraphs $content $slideSpec.items 18
    if ($slideSpec.intro) {
        [void]$content.TextFrame.TextRange.InsertBefore(([string]$slideSpec.intro + "`r"))
        $content.TextFrame.TextRange.Paragraphs(1).Font.Size = 16
        $content.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = $Palette.Muted
        $content.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Bullet.Visible = 0
    }
    Add-Fade $slide $content 1
    Add-Notes $slide $slideSpec
}

function Add-ContentSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $content = Find-Placeholder $slide @('Content') @(2,7)
    Set-BodyParagraphs $content $slideSpec.bullets 16
    Add-Fade $slide $content 1
    Add-Notes $slide $slideSpec
}

function Add-TwoContentSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Two Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $contentPlaceholders = Get-ContentPlaceholders $slide
    $left = if ($contentPlaceholders.Count -ge 1) { $contentPlaceholders[0] } else { $null }
    $right = if ($contentPlaceholders.Count -ge 2) { $contentPlaceholders[1] } else { $null }
    if ($left -eq $null) {
        $left = Add-TransparentTextBox $slide '' 48 126 408 366 18 $false $Palette.Ink -AllowEmpty
    }
    if ($right -eq $null) {
        $right = Add-TransparentTextBox $slide '' 488 126 408 366 18 $false $Palette.Ink -AllowEmpty
    }
    $leftText = @([string]$slideSpec.leftTitle) + @($slideSpec.leftBullets)
    $rightText = @([string]$slideSpec.rightTitle) + @($slideSpec.rightBullets)
    Set-BodyParagraphs $left $leftText 16
    Set-BodyParagraphs $right $rightText 16
    $left.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    $right.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    Add-Fade $slide $left 1
    Add-Fade $slide $right 1
    Add-Notes $slide $slideSpec
}

function Add-WhatToExpectSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0

    $outer = $slide.Shapes.AddShape(5, 54, 120, 850, 390)
    $outer.Fill.ForeColor.RGB = RgbLong 255 255 255
    $outer.Fill.Transparency = 1
    $outer.Line.ForeColor.RGB = $Palette.Orange
    $outer.Line.Weight = 0.9
    Set-LiPatternTag $outer 'WhatToExpect'

    $intro = [string]$slideSpec.intro
    if ($intro) {
        $introBox = Add-TransparentTextBox $slide $intro 78 142 790 46 16 $false $Palette.Ink
        Set-LiPatternTag $introBox 'WhatToExpect'
    }

    $sections = @($slideSpec.sections)
    if ($sections.Count -eq 0) {
        $sections = @($slideSpec.items)
    }
    if ($sections.Count -eq 0) {
        $sections = @(
            [pscustomobject]@{ label = 'Best Practices'; detail = 'Concepts and standards to reuse.'; role = 'best-practice' },
            [pscustomobject]@{ label = 'Why should I do this?'; detail = 'Motivation and risk context.'; role = 'why' },
            [pscustomobject]@{ label = 'How do I do this?'; detail = 'Implementation steps and examples.'; role = 'how' },
            [pscustomobject]@{ label = 'Demo or Exercise'; detail = 'Practice with a concrete artifact.'; role = 'demo' },
            [pscustomobject]@{ label = 'Discussion'; detail = 'Open questions and decisions.'; role = 'discussion' }
        )
    }

    $roleColors = @{
        'best-practice' = @{ Fill = RgbLong 244 238 249; Line = RgbLong 112 48 160 }
        'extra-note' = @{ Fill = RgbLong 248 242 252; Line = RgbLong 180 130 218 }
        'why' = @{ Fill = $Palette.LightYellow; Line = $Palette.Orange }
        'how' = @{ Fill = $Palette.LightGreen; Line = $Palette.Green }
        'what' = @{ Fill = $Palette.LightBlue; Line = $Palette.Blue }
        'demo' = @{ Fill = $Palette.LightCyan; Line = $Palette.Cyan }
        'discussion' = @{ Fill = $Palette.LightRed; Line = $Palette.Red }
    }
    $bandShapes = @()
    $y = 205
    foreach ($section in $sections) {
        $label = if ($section -is [string]) { [string]$section } else { [string]$section.label }
        $detail = if ($section -is [string]) { '' } else { [string]$section.detail }
        $role = if ($section -is [string]) { '' } else { [string]$section.role }
        $roleColor = if ($roleColors.ContainsKey($role)) { $roleColors[$role] } else { @{ Fill = $LightFills[$bandShapes.Count % $LightFills.Count]; Line = $AccentLines[$bandShapes.Count % $AccentLines.Count] } }
        $height = if ($detail) { 46 } else { 30 }
        $band = $slide.Shapes.AddShape(5, 78, $y, 650, $height)
        $band.Fill.ForeColor.RGB = $roleColor.Fill
        $band.Fill.Transparency = 0.02
        $band.Line.ForeColor.RGB = $roleColor.Line
        $band.Line.Weight = 0.9
        $text = if ($detail) { "$label`r$detail" } else { $label }
        $bandTextColor = $Palette.Ink
        Add-TextBoxInShape $band $text 12.5 $true $bandTextColor
        if ($detail) {
            $band.TextFrame.TextRange.Paragraphs(2).Font.Size = 10.5
            $band.TextFrame.TextRange.Paragraphs(2).Font.Bold = 0
        }
        $bandShapes += $band
        $y += $height + 10
    }
    Add-Fade $slide $outer 1
    $outer.Line.Visible = -1
    $outer.Line.ForeColor.RGB = $Palette.Orange
    $outer.Line.Weight = 0.9
    $group = Add-PatternGroup $slide $bandShapes 'WhatToExpect' 'Li_WhatToExpectBands'
    if ($group -ne $null) { Add-Fade $slide $group 1 }
    Add-Notes $slide $slideSpec
}

function Add-ProgressSidebarSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $items = @(Get-LocatorItems $slideSpec)
    if ($items.Count -eq 0 -and $slideSpec.steps) { $items = @($slideSpec.steps) }
    Add-ProgressRail $slide $items ([string]$slideSpec.activeProgressLabel) 54 145 230 28 | Out-Null

    $mainTitle = if ($slideSpec.mainTitle) { [string]$slideSpec.mainTitle } else { [string]$slideSpec.activeProgressLabel }
    if ($mainTitle) {
        $header = $slide.Shapes.AddShape(5, 330, 150, 420, 52)
        $header.Fill.ForeColor.RGB = $Palette.LightGreen
        $header.Line.ForeColor.RGB = $Palette.Green
        $header.Line.Weight = 1.5
        Add-TextBoxInShape $header $mainTitle 20 $true $Palette.Ink
        Set-LiPatternTag $header 'ProgressSidebar'
    }

    $cards = @($slideSpec.cards)
    if ($cards.Count -eq 0) { $cards = @($slideSpec.items) }
    if ($cards.Count -eq 0 -and $slideSpec.intro) { $cards = @([pscustomobject]@{ label = [string]$slideSpec.intro; role = 'highlight' }) }
    $x = 330
    $y = 235
    $shapes = @()
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $card = $cards[$i]
        $label = if ($card -is [string]) { [string]$card } else { [string]$card.label }
        $detail = if ($card -is [string]) { '' } else { [string]$card.detail }
        $role = if ($card -is [string]) { '' } else { [string]$card.role }
        $colors = Get-SemanticColors $role $i
        $cardText = if ($detail) { "$label`r$detail" } else { $label }
        $shape = Add-RoundedCard $slide $cardText ($x + (($i % 2) * 265)) ($y + ([Math]::Floor($i / 2) * 96)) 235 70 $i 14 -NoAnimation
        $shape.Fill.ForeColor.RGB = $colors.Fill
        $shape.Line.ForeColor.RGB = $colors.Line
        Set-LiPatternTag $shape 'ProgressSidebar'
        $shapes += $shape
    }
    $group = Add-PatternGroup $slide $shapes 'ProgressSidebar' 'Li_ProgressSidebarMain'
    if ($group -ne $null) { Add-Fade $slide $group 1 }
    Add-Notes $slide $slideSpec
}

function Add-ImageEvidenceSlide($presentation, $slideSpec, [string]$patternName) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    if ($slideSpec.intro) {
        $intro = Add-TransparentTextBox $slide ([string]$slideSpec.intro) 78 135 780 34 15 $false $Palette.Muted
        Set-LiPatternTag $intro $patternName
    }

    $images = @($slideSpec.images)
    $count = [Math]::Max(1, $images.Count)
    $shapes = @()
    for ($i = 0; $i -lt $images.Count; $i++) {
        $image = $images[$i]
        $path = if ($image -is [string]) { [string]$image } else { [string]$image.path }
        $caption = if ($image -is [string]) { '' } else { [string]$image.caption }
        $resolved = Resolve-AssetPath $path
        $left = if ($count -eq 1) { 78 } else { 78 + ($i * 405) }
        $top = 185
        $width = if ($count -eq 1) { 560 } else { 360 }
        $height = if ($count -eq 1) { 285 } else { 230 }
        if ($resolved -and (Test-Path -LiteralPath $resolved)) {
            $pic = $slide.Shapes.AddPicture($resolved, 0, -1, $left, $top, $width, $height)
            $pic.Line.ForeColor.RGB = $Palette.Muted
            $pic.Line.Weight = 0.75
            Set-LiPatternTag $pic $patternName
            $shapes += $pic
        } else {
            $placeholderText = if ($caption) { "Placeholder`r$caption" } else { "Placeholder`rAdd screenshot or evidence asset" }
            $pic = Add-RoundedCard $slide $placeholderText $left $top $width $height $i 16 -NoAnimation
            $pic.Fill.ForeColor.RGB = $Palette.LightGray
            $pic.Line.ForeColor.RGB = $Palette.Muted
            Set-LiPatternTag $pic $patternName
            $shapes += $pic
        }
        if ($caption) {
            $captionBox = Add-TransparentTextBox $slide $caption $left ($top + $height + 8) $width 26 11 $false $Palette.Muted
            Set-LiPatternTag $captionBox $patternName
            $shapes += $captionBox
        }
    }

    $callouts = @($slideSpec.callouts)
    for ($i = 0; $i -lt $callouts.Count; $i++) {
        $callout = $callouts[$i]
        $label = if ($callout -is [string]) { [string]$callout } else { [string]$callout.label }
        $role = if ($callout -is [string]) { '' } else { [string]$callout.role }
        $colors = Get-SemanticColors $role $i
        $box = Add-RoundedCard $slide $label 670 (190 + $i * 72) 220 54 $i 12.5 -NoAnimation
        $box.Fill.ForeColor.RGB = $colors.Fill
        $box.Line.ForeColor.RGB = $colors.Line
        Set-LiPatternTag $box $patternName
        $shapes += $box
    }
    $group = Add-PatternGroup $slide $shapes $patternName "Li_$patternName"
    if ($group -ne $null) { Add-Fade $slide $group 1 }
    Add-Notes $slide $slideSpec
}

function Get-StatusFill([string]$text) {
    if (-not $text) { return $null }
    if ($text -match '(?i)\b(pass|passed|complete|completed|validated|done|ok|ready)\b') { return $Palette.LightGreen }
    if ($text -match '(?i)\b(risk|gap|fail|failed|blocker|missing|open)\b') { return $Palette.LightRed }
    if ($text -match '(?i)\b(review|pending|next|watch|in progress|partial)\b') { return $Palette.LightYellow }
    return $null
}

function Add-ComparisonTableSlide($presentation, $slideSpec, [string]$patternName) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    if ($slideSpec.intro) {
        $intro = Add-TransparentTextBox $slide ([string]$slideSpec.intro) 78 135 790 34 15 $false $Palette.Muted
        Set-LiPatternTag $intro $patternName
    }
    $headers = @($slideSpec.headers)
    if ($headers.Count -eq 0) { $headers = @('Item', 'Evidence', 'Status', 'Next') }
    $rows = @($slideSpec.rows)
    if ($rows.Count -eq 0) { $rows = @($slideSpec.items) }
    $rowCount = [Math]::Max(2, $rows.Count + 1)
    $colCount = [Math]::Max(2, $headers.Count)
    $tableShape = $slide.Shapes.AddTable($rowCount, $colCount, 70, 190, 820, 255)
    Set-LiPatternTag $tableShape $patternName
    $table = $tableShape.Table
    for ($c = 1; $c -le $colCount; $c++) {
        $text = if ($c -le $headers.Count) { [string]$headers[$c - 1] } else { '' }
        $cell = $table.Cell(1, $c)
        $cell.Shape.Fill.ForeColor.RGB = $Palette.LightBlue
        $cell.Shape.TextFrame.TextRange.Text = $text
        $cell.Shape.TextFrame.TextRange.Font.Bold = -1
        $cell.Shape.TextFrame.TextRange.Font.Size = 12
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = $Palette.DeepBlue
    }
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row = $rows[$r]
        for ($c = 1; $c -le $colCount; $c++) {
            $value = ''
            if ($row -is [string]) {
                $parts = ([string]$row).Split('|')
                if (($c - 1) -lt $parts.Count) { $value = $parts[$c - 1].Trim() }
            } elseif ($row.cells) {
                $cells = @($row.cells)
                if (($c - 1) -lt $cells.Count) { $value = [string]$cells[$c - 1] }
            } else {
                $propName = [string]$headers[$c - 1]
                $prop = $row.PSObject.Properties | Where-Object { $_.Name -ieq $propName } | Select-Object -First 1
                if ($prop -ne $null) { $value = [string]$prop.Value }
            }
            $cell = $table.Cell($r + 2, $c)
            $cell.Shape.Fill.ForeColor.RGB = RgbLong 255 255 255
            $cell.Shape.Fill.Transparency = 0
            $cell.Shape.TextFrame.TextRange.Text = $value
            $cell.Shape.TextFrame.TextRange.Font.Size = 11
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = $Palette.Ink
            $statusFill = Get-StatusFill $value
            $headerText = if (($c - 1) -lt $headers.Count) { [string]$headers[$c - 1] } else { '' }
            $isStatusColumn = ($headerText -match '(?i)status|result|outcome')
            if ($isStatusColumn -and $statusFill -ne $null) {
                $cell.Shape.Fill.ForeColor.RGB = $statusFill
            } elseif (($r % 2) -eq 1) {
                $cell.Shape.Fill.ForeColor.RGB = $Palette.LightGray
                $cell.Shape.Fill.Transparency = 0.35
            }
        }
    }
    if ($slideSpec.callout) {
        $callout = Add-RoundedCard $slide ([string]$slideSpec.callout) 120 465 720 36 2 14 -NoAnimation
        $callout.Fill.ForeColor.RGB = $Palette.LightYellow
        $callout.Line.ForeColor.RGB = $Palette.Orange
        Set-LiPatternTag $callout $patternName
        Add-Fade $slide $callout 1
    }
    Add-Fade $slide $tableShape 1
    Add-Notes $slide $slideSpec
}

function Add-ConceptArtifactSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0

    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null -and $slideSpec.intro) {
        Set-Text $content ([string]$slideSpec.intro) 15 $false $Palette.Muted
        try { $content.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0 } catch {}
        Set-LiPatternTag $content 'ConceptArtifact'
    }

    $shapes = @()
    $concepts = @($slideSpec.concepts)
    if ($concepts.Count -eq 0) { $concepts = @($slideSpec.items) }
    for ($i = 0; $i -lt $concepts.Count; $i++) {
        $item = $concepts[$i]
        $label = if ($item -is [string]) { [string]$item } else { [string]$item.label }
        $detail = if ($item -is [string]) { '' } else { [string]$item.detail }
        $role = if ($item -is [string]) { 'model' } else { [string]$item.role }
        $colors = Get-SemanticColors $role $i
        $text = if ($detail) { "$label`r$detail" } else { $label }
        $card = Add-RoundedCard $slide $text 68 (168 + $i * 72) 360 56 $i 14 -NoAnimation
        $card.Fill.ForeColor.RGB = $colors.Fill
        $card.Line.ForeColor.RGB = $colors.Line
        Set-LiPatternTag $card 'ConceptArtifact'
        if ($detail) {
            $card.TextFrame.TextRange.Paragraphs(1).Font.Size = 15
            $card.TextFrame.TextRange.Paragraphs(2).Font.Size = 11.5
            $card.TextFrame.TextRange.Paragraphs(2).Font.Bold = 0
            $card.TextFrame.TextRange.Paragraphs(2).Font.Color.RGB = $Palette.Muted
        }
        $shapes += $card
    }

    $images = @($slideSpec.images)
    $artifactLabel = if ($slideSpec.artifactLabel) { [string]$slideSpec.artifactLabel } else { 'Artifact or evidence' }
    $artifactShape = $null
    if ($images.Count -gt 0) {
        $image = $images[0]
        $path = if ($image -is [string]) { [string]$image } else { [string]$image.path }
        $resolved = Resolve-AssetPath $path
        if ($resolved -and (Test-Path -LiteralPath $resolved)) {
            $artifactShape = $slide.Shapes.AddPicture($resolved, 0, -1, 486, 158, 360, 232)
            $artifactShape.Line.ForeColor.RGB = $Palette.Muted
            $artifactShape.Line.Weight = 0.75
        }
    }
    if ($artifactShape -eq $null) {
        $artifactShape = Add-RoundedCard $slide "Placeholder`r$artifactLabel" 486 158 360 232 1 17 -NoAnimation
        $artifactShape.Fill.ForeColor.RGB = $Palette.LightGray
        $artifactShape.Line.ForeColor.RGB = $Palette.Muted
    }
    Set-LiPatternTag $artifactShape 'ConceptArtifact'
    $shapes += $artifactShape

    $callouts = @($slideSpec.callouts)
    for ($i = 0; $i -lt $callouts.Count; $i++) {
        $callout = $callouts[$i]
        $label = if ($callout -is [string]) { [string]$callout } else { [string]$callout.label }
        $role = if ($callout -is [string]) { 'highlight' } else { [string]$callout.role }
        $colors = Get-SemanticColors $role $i
        $box = Add-RoundedCard $slide $label (486 + ($i % 2) * 184) (410 + [Math]::Floor($i / 2) * 48) 170 38 $i 11.5 -NoAnimation
        $box.Fill.ForeColor.RGB = $colors.Fill
        $box.Line.ForeColor.RGB = $colors.Line
        Set-LiPatternTag $box 'ConceptArtifact'
        $shapes += $box
    }

    $group = Add-PatternGroup $slide $shapes 'ConceptArtifact' 'Li_ConceptArtifact'
    if ($group -ne $null) { Add-Fade $slide $group 1 }
    Add-Notes $slide $slideSpec
}

function Add-VModelToolMapSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null -and $slideSpec.intro) {
        Set-Text $content ([string]$slideSpec.intro) 15 $false $Palette.Muted
        try { $content.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0 } catch {}
        Set-LiPatternTag $content 'VModelToolMap'
    }

    $phases = @($slideSpec.phases)
    if ($phases.Count -eq 0) { $phases = @($slideSpec.steps) }
    if ($phases.Count -eq 0) {
        $phases = @('Scope', 'Architecture', 'Model', 'Test', 'Deploy')
    }
    $positions = @(
        @{ X = 78; Y = 185 },
        @{ X = 238; Y = 276 },
        @{ X = 410; Y = 354 },
        @{ X = 582; Y = 276 },
        @{ X = 742; Y = 185 }
    )
    $phaseShapes = @()
    $phaseCards = @()
    for ($i = 0; $i -lt [Math]::Min($phases.Count, $positions.Count); $i++) {
        $phase = $phases[$i]
        $label = if ($phase -is [string]) { [string]$phase } else { [string]$phase.label }
        $tool = if ($phase -is [string]) { '' } else { [string]$phase.tool }
        $pos = $positions[$i]
        $card = Add-RoundedCard $slide $label $pos.X $pos.Y 136 54 $i 13.5 -NoAnimation
        $card.Fill.ForeColor.RGB = $Palette.LightBlue
        $card.Line.ForeColor.RGB = $Palette.Blue
        Set-LiPatternTag $card 'VModelToolMap'
        $phaseShapes += $card
        $phaseCards += $card
        if ($tool) {
            $toolBox = Add-RoundedCard $slide $tool ($pos.X + 12) ($pos.Y + 62) 112 24 $i 9.5 -NoAnimation
            $toolBox.Fill.ForeColor.RGB = $Palette.LightYellow
            $toolBox.Line.ForeColor.RGB = $Palette.Orange
            Set-LiPatternTag $toolBox 'VModelToolMap'
            $phaseShapes += $toolBox
        }
    }
    for ($i = 0; $i -lt ($phaseCards.Count - 1); $i++) {
        try {
            $a = $phaseCards[$i]
            $b = $phaseCards[$i + 1]
            $line = $slide.Shapes.AddConnector(1, ($a.Left + $a.Width), ($a.Top + 27), $b.Left, ($b.Top + 27))
            $line.Line.ForeColor.RGB = $Palette.Orange
            $line.Line.Weight = 1.2
            Set-LiPatternTag $line 'VModelToolMap'
            Add-Wipe $slide $line 1 1
        } catch {}
    }

    $callouts = @($slideSpec.callouts)
    for ($i = 0; $i -lt $callouts.Count; $i++) {
        $label = if ($callouts[$i] -is [string]) { [string]$callouts[$i] } else { [string]$callouts[$i].label }
        $box = Add-RoundedCard $slide $label 110 (438 + $i * 35) 700 28 $i 11.5 -NoAnimation
        $box.Fill.ForeColor.RGB = $Palette.LightGreen
        $box.Line.ForeColor.RGB = $Palette.Green
        Set-LiPatternTag $box 'VModelToolMap'
        $phaseShapes += $box
    }

    foreach ($shape in $phaseShapes) { Add-Fade $slide $shape 1 }
    Add-Notes $slide $slideSpec
}

function Add-ProcessStateDiagramSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null -and $slideSpec.intro) {
        Set-Text $content ([string]$slideSpec.intro) 15 $false $Palette.Muted
        try { $content.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0 } catch {}
        Set-LiPatternTag $content 'ProcessStateDiagram'
    }

    $steps = @($slideSpec.steps)
    $count = [Math]::Max(1, $steps.Count)
    $width = [Math]::Min(150, (812 - (($count - 1) * 18)) / $count)
    $left = 72
    $top = 230
    $shapes = @()
    for ($i = 0; $i -lt $count; $i++) {
        $step = $steps[$i]
        $label = if ($step -is [string]) { [string]$step } else { [string]$step.label }
        $status = if ($step -is [string]) { '' } else { [string]$step.status }
        $detail = if ($step -is [string]) { '' } else { [string]$step.detail }
        $colors = Get-SemanticColors $status $i
        $x = $left + ($i * ($width + 18))
        $cardText = if ($detail) { "$label`r$detail" } else { $label }
        $card = Add-RoundedCard $slide $cardText $x $top $width 92 $i 13 -NoAnimation
        $card.Fill.ForeColor.RGB = $colors.Fill
        $card.Line.ForeColor.RGB = $colors.Line
        Set-LiPatternTag $card 'ProcessStateDiagram'
        $shapes += $card
        if ($status) {
            $mark = Add-RoundedCard $slide $status ($x + 10) ($top - 32) 64 24 $i 9.5 -NoAnimation
            $mark.Fill.ForeColor.RGB = $colors.Fill
            $mark.Line.ForeColor.RGB = $colors.Line
            Set-LiPatternTag $mark 'ProcessStateDiagram'
            $shapes += $mark
        }
        if ($i -lt ($count - 1)) {
            $arrow = Add-Arrow $slide ($x + $width + 3) ($top + 34) 20 18 -NoAnimation
            Set-LiPatternTag $arrow 'ProcessStateDiagram'
            $shapes += $arrow
        }
    }
    foreach ($shape in $shapes) { Add-Fade $slide $shape 1 }
    Add-Notes $slide $slideSpec
}

function Add-DemoExerciseSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $panels = @($slideSpec.panels)
    if ($panels.Count -eq 0) {
        $panels = @(
            [pscustomobject]@{ label = 'Start'; detail = [string]$slideSpec.start; role = 'current' },
            [pscustomobject]@{ label = 'Do'; detail = [string]$slideSpec.task; role = 'action' },
            [pscustomobject]@{ label = 'Keep'; detail = [string]$slideSpec.output; role = 'verification' }
        )
    }
    $shapes = @()
    for ($i = 0; $i -lt $panels.Count; $i++) {
        $panel = $panels[$i]
        $label = [string]$panel.label
        $detail = [string]$panel.detail
        $role = [string]$panel.role
        $colors = Get-SemanticColors $role $i
        $panelText = if ($detail) { "$label`r$detail" } else { $label }
        $shape = Add-RoundedCard $slide $panelText (84 + $i * 270) 210 235 132 $i 16 -NoAnimation
        $shape.Fill.ForeColor.RGB = if ($role -match 'demo|exercise') { $Palette.LightCyan } else { $colors.Fill }
        $shape.Line.ForeColor.RGB = if ($role -match 'demo|exercise') { $Palette.Cyan } else { $colors.Line }
        Set-LiPatternTag $shape 'DemoExercise'
        $shapes += $shape
    }
    $group = Add-PatternGroup $slide $shapes 'DemoExercise' 'Li_DemoExercise'
    if ($group -ne $null) { Add-Fade $slide $group 1 }
    Add-Notes $slide $slideSpec
}

function Add-RecapSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    if ($slideSpec.subtitle) {
        $subtitle = Add-TransparentTextBox $slide ([string]$slideSpec.subtitle) 78 105 790 36 19 $false $Palette.Blue
        Set-LiPatternTag $subtitle 'Recap'
    }
    $items = @($slideSpec.items)
    if ($items.Count -eq 0) { $items = @($slideSpec.bullets) }
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null) {
        Set-BodyParagraphs $content $items 16
        Set-LiPatternTag $content 'Recap'
        Add-Fade $slide $content 1
    }
    Add-Notes $slide $slideSpec
}

function Add-SectionSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Section Header'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    if ($slideSpec.outcome) {
        $outcome = Add-TransparentTextBox $slide ([string]$slideSpec.outcome) 220 270 520 42 16 $false $Palette.Muted
        Set-LiPatternTag $outcome 'SectionOutcome'
    }
    $items = @(Get-LocatorItems $slideSpec)
    if ($items.Count -gt 0) {
        $count = $items.Count
        if ($count -gt 0) {
            $startLeft = 140
            $top = 342
            $width = [Math]::Min(118, 720 / $count)
            $locatorShapes = @()
            for ($i = 0; $i -lt $count; $i++) {
                $label = if ($items[$i] -is [string]) { [string]$items[$i] } else { [string]$items[$i].label }
                $active = $false
                try { $active = [bool]$items[$i].active } catch {}
                $shape = $slide.Shapes.AddShape(5, ($startLeft + $i * ($width + 8)), $top, $width, 36)
                $shape.Fill.ForeColor.RGB = if ($active) { $Palette.LightYellow } else { $Palette.LightGray }
                $shape.Fill.Transparency = if ($active) { 0.02 } else { 0.12 }
                $shape.Line.ForeColor.RGB = if ($active) { $Palette.Orange } else { $Palette.Muted }
                $shape.Line.Weight = if ($active) { 1.5 } else { 0.75 }
                $shape.TextFrame.TextRange.Text = $label
                $shape.TextFrame.TextRange.Font.Size = 11
                $shape.TextFrame.TextRange.Font.Bold = if ($active) { -1 } else { 0 }
                $shape.TextFrame.TextRange.Font.Color.RGB = $Palette.Ink
                $shape.TextFrame.MarginLeft = 4
                $shape.TextFrame.MarginRight = 4
                $shape.TextFrame.MarginTop = 3
                $shape.TextFrame.MarginBottom = 3
                $locatorShapes += $shape
            }
            Add-PatternGroup $slide $locatorShapes 'SectionLocator' 'Li_SectionLocator' | Out-Null
        }
    }
    Add-Notes $slide $slideSpec
}

function Add-ProcessSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $variant = if ($slideSpec.variant) { [string]$slideSpec.variant } else { 'horizontal-cards' }
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null) {
        Set-Text $content ([string]$slideSpec.intro) 16 $false $Palette.Muted
        Set-LiPatternTag $content "Process_$variant"
        try { $content.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0 } catch {}
    }

    $steps = @($slideSpec.steps)
    $count = [Math]::Max(1, $steps.Count)
    $left = 54
    $top = 220
    $gap = 22
    $width = [Math]::Min(150, (830 - (($count - 1) * $gap)) / $count)
    if ($width -lt 118) { $width = 118 }
    $cards = @()
    $arrows = @()
    $xPositions = @()
    for ($i = 0; $i -lt $count; $i++) {
        $x = $left + ($i * ($width + $gap))
        $xPositions += $x
        $label = if ($steps[$i] -is [string]) { [string]$steps[$i] } else { [string]$steps[$i].label }
        $detail = if ($steps[$i] -is [string]) { '' } else { [string]$steps[$i].detail }
        $text = if ($detail) { "$label`r$detail" } else { $label }
        $card = Add-RoundedCard $slide $text $x $top $width 94 $i 15 -NoAnimation
        Set-LiPatternTag $card "Process_$variant"
        if ($detail) {
            $card.TextFrame.TextRange.Paragraphs(1).Font.Size = 17
            $card.TextFrame.TextRange.Paragraphs(2).Font.Size = 12.5
            $card.TextFrame.TextRange.Paragraphs(2).Font.Bold = 0
            $card.TextFrame.TextRange.Paragraphs(2).Font.Color.RGB = $Palette.Muted
        }
        $cards += $card
    }
    for ($i = 0; $i -lt ($count - 1); $i++) {
        $arrow = Add-Arrow $slide ($xPositions[$i] + $width + 4) ($top + 32) 24 20 -NoAnimation
        Set-LiPatternTag $arrow "Process_$variant"
        $arrows += $arrow
    }
    if ($cards.Count -gt 0) {
        Add-Fade $slide $cards[0] 1
    }
    for ($i = 1; $i -lt $cards.Count; $i++) {
        $group = Group-Shapes $slide @($arrows[$i - 1], $cards[$i]) "Build_$($i + 1)"
        if ($group -ne $null) {
            Set-LiPatternTag $group "Process_$variant"
            Add-Fade $slide $group 1
        } else {
            Add-Fade $slide $arrows[$i - 1] 1
            Add-Fade $slide $cards[$i] 1
        }
    }
    Add-Notes $slide $slideSpec
}

function Add-DecisionSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null) {
        Set-Text $content ([string]$slideSpec.intro) 16 $false $Palette.Muted
        try { $content.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0 } catch {}
    }
    $items = @($slideSpec.items)
    $count = [Math]::Max(1, $items.Count)
    $width = [Math]::Min(250, (830 - (($count - 1) * 28)) / $count)
    $left = 58
    for ($i = 0; $i -lt $count; $i++) {
        $text = if ($items[$i] -is [string]) { [string]$items[$i] } else { [string]$items[$i].label }
        Add-RoundedCard $slide $text ($left + $i * ($width + 28)) 225 $width 118 $i 18 | Out-Null
    }
    Add-Notes $slide $slideSpec
}

function Add-ArtifactMapSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null) {
        Set-Text $content ([string]$slideSpec.intro) 15 $false $Palette.Muted
        try { $content.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0 } catch {}
    }
    $artifacts = @($slideSpec.artifacts)
    $left = 60
    $top = 185
    $groupShapes = @()
    for ($i = 0; $i -lt $artifacts.Count; $i++) {
        $artifact = $artifacts[$i]
        $label = [string]$artifact.label
        $detail = [string]$artifact.detail
        $role = [string]$artifact.role
        $x = $left + (($i % 3) * 280)
        $y = $top + ([Math]::Floor($i / 3) * 118)
        $colors = Get-SemanticColors $role $i
        $cardText = if ($detail) { "$label`r$detail" } else { $label }
        $card = Add-RoundedCard $slide $cardText $x $y 238 82 $i 15 -NoAnimation
        $card.Fill.ForeColor.RGB = $colors.Fill
        $card.Line.ForeColor.RGB = $colors.Line
        Set-LiPatternTag $card 'ArtifactMap'
        if ($detail) {
            $card.TextFrame.TextRange.Paragraphs(1).Font.Size = 16
            $card.TextFrame.TextRange.Paragraphs(2).Font.Size = 11.5
            $card.TextFrame.TextRange.Paragraphs(2).Font.Bold = 0
            $card.TextFrame.TextRange.Paragraphs(2).Font.Color.RGB = $Palette.Muted
        }
        $rolePill = $slide.Shapes.AddShape(5, ($x + 14), ($y + 55), 124, 18)
        $rolePill.Fill.ForeColor.RGB = $colors.Line
        $rolePill.Fill.Transparency = 0.05
        $rolePill.Line.Visible = 0
        Add-TextBoxInShape $rolePill $role 8.5 $true (RgbLong 255 255 255)
        Set-LiPatternTag $rolePill 'ArtifactMap'
        $group = Add-PatternGroup $slide @($card, $rolePill) 'ArtifactMap' "Li_Artifact_$($i + 1)"
        if ($group -ne $null) { Add-Fade $slide $group 1 } else { Add-Fade $slide $card 1 }
        $groupShapes += $group
    }
    Add-Notes $slide $slideSpec
}

function Add-CodeReviewSlide($presentation, $slideSpec) {
    $slide = $presentation.Slides.AddSlide($presentation.Slides.Count + 1, (Get-Layout $presentation 'Title and Content'))
    Set-Text (Find-Placeholder $slide @('Title') @(1,3)) ([string]$slideSpec.title) 0 $false 0
    $content = Find-Placeholder $slide @('Content') @(2,7)
    if ($content -ne $null) {
        Set-Text $content ([string]$slideSpec.intro) 15 $false $Palette.Muted
        try { $content.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0 } catch {}
    }
    $code = [string]$slideSpec.code
    $codeBox = $slide.Shapes.AddShape(5, 58, 180, 510, 230)
    $codeBox.Fill.ForeColor.RGB = $Palette.LightGray
    $codeBox.Fill.Transparency = 0.04
    $codeBox.Line.ForeColor.RGB = $Palette.Muted
    $codeBox.TextFrame.TextRange.Text = $code
    $codeBox.TextFrame.TextRange.Font.Name = 'Consolas'
    $codeBox.TextFrame.TextRange.Font.Size = 12
    $codeBox.TextFrame.TextRange.Font.Color.RGB = $Palette.Ink
    $codeBox.TextFrame.MarginLeft = 12
    $codeBox.TextFrame.MarginTop = 8
    $notes = @($slideSpec.callouts)
    for ($i = 0; $i -lt $notes.Count; $i++) {
        $callout = if ($notes[$i] -is [string]) { [string]$notes[$i] } else { [string]$notes[$i].label }
        Add-RoundedCard $slide $callout 600 (180 + $i * 76) 240 58 $i 13 | Out-Null
    }
    Add-Fade $slide $codeBox 1
    Add-Notes $slide $slideSpec
}

function Join-MarkdownList($items) {
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($items)) {
        if ($null -eq $item) { continue }
        if ($item -is [string]) {
            if ($item.Trim()) { $lines.Add("- $item") | Out-Null }
        } elseif ($item.label) {
            $detail = [string]$item.detail
            $line = "- $($item.label)"
            if ($detail) { $line += ": $detail" }
            $lines.Add($line) | Out-Null
        } else {
            $lines.Add("- $($item | ConvertTo-Json -Compress -Depth 4)") | Out-Null
        }
    }
    return ($lines -join "`r`n")
}

function Write-PresentationMarkdownArtifacts($specNode, [string]$deckPath, [string]$templateKind, [string]$groundTruthPath, [string]$speechPath) {
    $deckDir = Split-Path -Parent $deckPath
    if (-not $groundTruthPath) { $groundTruthPath = Join-Path $deckDir 'ground-truth.md' }
    if (-not $speechPath) { $speechPath = Join-Path $deckDir 'speech.md' }

    $resolvedGroundTruth = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($groundTruthPath)
    $resolvedSpeech = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($speechPath)
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resolvedGroundTruth) | Out-Null
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $resolvedSpeech) | Out-Null

    $gt = New-Object System.Collections.Generic.List[string]
    $gt.Add("# Ground Truth") | Out-Null
    $gt.Add("") | Out-Null
    $gt.Add("- Title: $($specNode.title)") | Out-Null
    $gt.Add("- Subtitle: $($specNode.subtitle)") | Out-Null
    $gt.Add("- Use case: $($specNode.useCase)") | Out-Null
    $gt.Add("- Audience: $($specNode.audience)") | Out-Null
    $gt.Add("- Purpose: $($specNode.purpose)") | Out-Null
    $gt.Add("- Template: $templateKind") | Out-Null
    $gt.Add("- Output: $deckPath") | Out-Null
    $gt.Add("- Status: generated from JSON spec; update this ledger when content, assets, or slide order changes.") | Out-Null
    $gt.Add("") | Out-Null
    $gt.Add("## Source Facts") | Out-Null
    $facts = Join-MarkdownList $specNode.sourceFacts
    $gt.Add($(if ($facts) { $facts } else { "- Source facts not supplied in spec." })) | Out-Null
    $gt.Add("") | Out-Null
    $gt.Add("## Progress Labels") | Out-Null
    $progress = Join-MarkdownList $specNode.progressLabels
    $gt.Add($(if ($progress) { $progress } else { "- No shared progress labels supplied." })) | Out-Null
    $gt.Add("") | Out-Null
    $gt.Add("## Slide Ledger") | Out-Null
    $gt.Add("| Slide | Type | Title | Intent | Assets / gaps |") | Out-Null
    $gt.Add("| --- | --- | --- | --- | --- |") | Out-Null
    $gt.Add("| 1 | cover | $($specNode.title) | Establish title, ownership, and delivery context. | Template cover. |") | Out-Null
    $slideNumber = 2
    foreach ($slideSpec in @($specNode.slides)) {
        $missing = Join-MarkdownList $slideSpec.missingInputs
        $assetStatus = [string]$slideSpec.assetStatus
        $placeholderStatus = [string]$slideSpec.placeholderStatus
        $gapText = @($assetStatus, $placeholderStatus, ($missing -replace "`r?`n", '<br>')) | Where-Object { $_ }
        if ($gapText.Count -eq 0) { $gapText = @('No open gap recorded.') }
        $reason = ([string]$slideSpec.reason) -replace '\|', '/'
        $title = ([string]$slideSpec.title) -replace '\|', '/'
        $gt.Add("| $slideNumber | $($slideSpec.type) | $title | $reason | $($gapText -join '<br>') |") | Out-Null
        $slideNumber++
    }
    $gt.Add("") | Out-Null
    $gt.Add("## Open Inputs") | Out-Null
    $openInputs = Join-MarkdownList $specNode.missingInputs
    $gt.Add($(if ($openInputs) { $openInputs } else { "- None recorded." })) | Out-Null
    Set-Content -LiteralPath $resolvedGroundTruth -Value ($gt -join "`r`n") -Encoding UTF8

    $speech = New-Object System.Collections.Generic.List[string]
    $speech.Add("# Speech") | Out-Null
    $speech.Add("") | Out-Null
    $speech.Add("## Slide 1 - $($specNode.title)") | Out-Null
    $coverTalk = if ($specNode.coverTalkTrack) { [string]$specNode.coverTalkTrack } else { "Introduce the topic, audience context, and what the audience should be able to do after the session." }
    $speech.Add($coverTalk) | Out-Null
    $slideNumber = 2
    foreach ($slideSpec in @($specNode.slides)) {
        $speech.Add("") | Out-Null
        $speech.Add("## Slide $slideNumber - $($slideSpec.title)") | Out-Null
        $speech.Add([string]$slideSpec.talkTrack) | Out-Null
        if ($slideSpec.builds) {
            $speech.Add("") | Out-Null
            $speech.Add("Build sequence:") | Out-Null
            foreach ($build in @($slideSpec.builds)) {
                $speech.Add("- $build") | Out-Null
            }
        }
        $slideNumber++
    }
    Set-Content -LiteralPath $resolvedSpeech -Value ($speech -join "`r`n") -Encoding UTF8

    return [pscustomobject]@{
        GroundTruthPath = $resolvedGroundTruth
        SpeechPath = $resolvedSpeech
    }
}

$powerPoint = $null
$presentation = $null
try {
    Test-VisibleText $spec 'spec'
    Test-Spec $spec

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($templatePath, 0, 0, -1)
    $resourceSlideCount = $presentation.Slides.Count

    $cover = $presentation.Slides.Item(1)
    if ($cover.CustomLayout.Name -ne 'Title Slide') {
        throw "Template first slide must use Title Slide layout."
    }
    Set-CoverText $cover $spec

    foreach ($slideSpec in @($spec.slides)) {
        switch ([string]$slideSpec.type) {
            'agenda' { Add-AgendaSlide $presentation $slideSpec }
            'learning-path' { Add-AgendaSlide $presentation $slideSpec }
            'section' { Add-SectionSlide $presentation $slideSpec }
            'chapter-divider' { Add-SectionSlide $presentation $slideSpec }
            'what-to-expect' { Add-WhatToExpectSlide $presentation $slideSpec }
            'process' { Add-ProcessSlide $presentation $slideSpec }
            'process-build-series' { Add-ProgressSidebarSlide $presentation $slideSpec }
            'progress-sidebar' { Add-ProgressSidebarSlide $presentation $slideSpec }
            'decision' { Add-DecisionSlide $presentation $slideSpec }
            'two-content' { Add-TwoContentSlide $presentation $slideSpec }
            'content' { Add-ContentSlide $presentation $slideSpec }
            'chapter-objectives' { Add-ContentSlide $presentation $slideSpec }
            'recap' { Add-RecapSlide $presentation $slideSpec }
            'recap-bridge' { Add-RecapSlide $presentation $slideSpec }
            'artifact-map' { Add-ArtifactMapSlide $presentation $slideSpec }
            'artifact-review' { Add-ArtifactMapSlide $presentation $slideSpec }
            'architecture-layer' { Add-ArtifactMapSlide $presentation $slideSpec }
            'concept-artifact' { Add-ConceptArtifactSlide $presentation $slideSpec }
            'v-model-tool-map' { Add-VModelToolMapSlide $presentation $slideSpec }
            'process-state-diagram' { Add-ProcessStateDiagramSlide $presentation $slideSpec }
            'code-review-excerpt' { Add-CodeReviewSlide $presentation $slideSpec }
            'code-to-model-review' { Add-CodeReviewSlide $presentation $slideSpec }
            'image-evidence' { Add-ImageEvidenceSlide $presentation $slideSpec 'ImageEvidence' }
            'screenshot-evidence' { Add-ImageEvidenceSlide $presentation $slideSpec 'ScreenshotEvidence' }
            'screenshot-callout' { Add-ImageEvidenceSlide $presentation $slideSpec 'ScreenshotCallout' }
            'model-screenshot' { Add-ImageEvidenceSlide $presentation $slideSpec 'ModelScreenshot' }
            'comparison-table' { Add-ComparisonTableSlide $presentation $slideSpec 'ComparisonTable' }
            'comparison-evidence-table' { Add-ComparisonTableSlide $presentation $slideSpec 'ComparisonEvidenceTable' }
            'results-table' { Add-ComparisonTableSlide $presentation $slideSpec 'ResultsTable' }
            'demo-exercise' { Add-DemoExerciseSlide $presentation $slideSpec }
            'exercise-demo' { Add-DemoExerciseSlide $presentation $slideSpec }
            default { throw "Unsupported slide type: $($slideSpec.type)" }
        }
    }

    for ($i = 2; $i -le $resourceSlideCount; $i++) {
        $presentation.Slides.Item(2).Delete()
    }

    Remove-EmptyTextShapes $presentation

    foreach ($slide in $presentation.Slides) {
        try { $slide.SlideShowTransition.EntryEffect = 0 } catch {}
    }

    $createdSlideCount = [int]$presentation.Slides.Count
    $presentation.SaveAs($resolvedOutput)
    $markdownArtifacts = Write-PresentationMarkdownArtifacts $spec $resolvedOutput $TemplateKind $GroundTruthPath $SpeechPath

    if ($ExportPreview) {
        try { $presentation.Close() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
        $presentation = $null
        try { $powerPoint.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null } catch {}
        $powerPoint = $null

        $previewDir = Join-Path $outputDir (([System.IO.Path]::GetFileNameWithoutExtension($resolvedOutput)) + '-previews')
        & (Join-Path $PSScriptRoot 'export_pptx_previews.ps1') -DeckPath $resolvedOutput -OutDir $previewDir -Slides 'all'
    }

    [pscustomobject]@{
        OutputPath = $resolvedOutput
        SlideCount = $createdSlideCount
        TemplateKind = $TemplateKind
        TemplatePath = $templatePath
        GroundTruthPath = $markdownArtifacts.GroundTruthPath
        SpeechPath = $markdownArtifacts.SpeechPath
    } | Format-List
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
