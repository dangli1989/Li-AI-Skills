param(
    [Parameter(Mandatory = $true)]
    [string]$DeckPath,

    [ValidateSet('public', 'confidential')]
    [string]$TemplateKind = 'public',

    [ValidateSet('training-workshop', 'customer-wrap-up', 'internal-sharing', 'general')]
    [string]$UseCase = 'general',

    [string]$PreviewDir,

    [string]$SpecPath,

    [string]$GroundTruthPath,

    [string]$SpeechPath,

    [switch]$Strict,

    [switch]$AsJson
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $DeckPath)) {
    throw "Deck not found: $DeckPath"
}

$skillRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$templatePath = Join-Path $skillRoot "assets\li-mathworks-presentation\pptx\$TemplateKind.pptx"
if (-not (Test-Path -LiteralPath $templatePath)) {
    throw "Template not found: $templatePath"
}

$allowedLayouts = @('Title Slide', 'Agenda', 'Section Header', 'Title and Content', 'Two Content', 'Feature', 'Title Only')
$errors = New-Object System.Collections.Generic.List[string]
$styleErrors = New-Object System.Collections.Generic.List[string]
$warnings = New-Object System.Collections.Generic.List[string]
$resolvedDeck = (Resolve-Path -LiteralPath $DeckPath).Path
$spec = $null

if ($SpecPath) {
    if (-not (Test-Path -LiteralPath $SpecPath)) {
        throw "Spec not found: $SpecPath"
    }
    $spec = Get-Content -Raw -LiteralPath $SpecPath | ConvertFrom-Json
}

function Get-ProgressItemsFromSpec($slideSpec, $progressLabels) {
    if ($slideSpec.locator) { return @($slideSpec.locator) }
    if ($slideSpec.activeProgressLabel -and $progressLabels.Count -gt 0) {
        $activeLabel = [string]$slideSpec.activeProgressLabel
        return @($progressLabels | ForEach-Object {
            [pscustomobject]@{
                label = [string]$_
                active = ([string]$_ -eq $activeLabel)
            }
        })
    }
    return @()
}

function Test-ReviewSpec($specNode, $styleErrors, $warnings, [string]$useCase, [bool]$strictMode) {
    if ($null -eq $specNode) { return }
    $progressLabels = @($specNode.progressLabels | ForEach-Object { [string]$_ })
    $progressSet = @{}
    foreach ($label in $progressLabels) {
        $key = $label.ToLowerInvariant()
        if ($progressSet.ContainsKey($key)) {
            $styleErrors.Add("Spec progressLabels contains duplicate label '$label'.") | Out-Null
        }
        $progressSet[$key] = $true
    }

    $types = @{}
    $slideNumber = 1
    foreach ($slideSpec in @($specNode.slides)) {
        $type = [string]$slideSpec.type
        if (-not $types.ContainsKey($type)) { $types[$type] = 0 }
        $types[$type]++

        if (-not $slideSpec.reason) {
            $styleErrors.Add("Spec slide $slideNumber is missing reason.") | Out-Null
        }
        if (-not $slideSpec.talkTrack) {
            $styleErrors.Add("Spec slide $slideNumber is missing talkTrack for speech.md.") | Out-Null
        }

        $locator = @(Get-ProgressItemsFromSpec $slideSpec $progressLabels)
        if ($locator.Count -gt 0) {
            $seen = @{}
            $activeCount = 0
            foreach ($item in $locator) {
                $label = if ($item -is [string]) { [string]$item } else { [string]$item.label }
                $key = $label.ToLowerInvariant()
                if ($seen.ContainsKey($key)) {
                    $styleErrors.Add("Spec slide $slideNumber has duplicate locator/progress label '$label'.") | Out-Null
                }
                $seen[$key] = $true
                try { if ([bool]$item.active) { $activeCount++ } } catch {}
            }
            if ($activeCount -ne 1) {
                $styleErrors.Add("Spec slide $slideNumber has $activeCount active locator/progress labels; expected exactly one.") | Out-Null
            }
        }

        if ($slideSpec.activeProgressLabel -and $progressLabels.Count -gt 0) {
            $activeKey = ([string]$slideSpec.activeProgressLabel).ToLowerInvariant()
            if (-not $progressSet.ContainsKey($activeKey)) {
                $styleErrors.Add("Spec slide $slideNumber activeProgressLabel '$($slideSpec.activeProgressLabel)' is not in progressLabels.") | Out-Null
            }
        }
        $slideNumber++
    }

    if ($useCase -eq 'training-workshop' -or $strictMode) {
        if (-not ($types.ContainsKey('what-to-expect'))) {
            $warnings.Add('Training/workshop spec has no what-to-expect page; add one unless this is a short excerpt deck.') | Out-Null
        }
        $evidenceTypes = @('image-evidence', 'screenshot-evidence', 'model-screenshot', 'screenshot-callout', 'comparison-table', 'comparison-evidence-table', 'results-table', 'artifact-map', 'artifact-review', 'concept-artifact', 'v-model-tool-map', 'process-state-diagram', 'code-review-excerpt', 'code-to-model-review')
        $evidenceCount = 0
        foreach ($evidenceType in $evidenceTypes) {
            if ($types.ContainsKey($evidenceType)) { $evidenceCount += $types[$evidenceType] }
        }
        if ($evidenceCount -lt 2) {
            $styleErrors.Add("Training/workshop spec has only $evidenceCount source-derived artifact/evidence slide types; add concrete code/model/screenshot/table evidence.") | Out-Null
        }
    }
}

function Get-LiPattern($shape) {
    $alt = ''
    try { $alt = [string]$shape.AlternativeText } catch {}
    if ($alt -match 'LiPattern:([^;\r\n]+)') { return $Matches[1] }
    $name = ''
    try { $name = [string]$shape.Name } catch {}
    if ($name -match '^Li_([^_]+)') { return $Matches[1] }
    return ''
}

Test-ReviewSpec $spec $styleErrors $warnings $UseCase ([bool]$Strict)

$artifactBaseDir = if ($SpecPath) { Split-Path -Parent (Resolve-Path -LiteralPath $SpecPath).Path } else { Split-Path -Parent $resolvedDeck }
$resolvedGroundTruthPath = if ($GroundTruthPath) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($GroundTruthPath) } else { Join-Path $artifactBaseDir 'ground-truth.md' }
$resolvedSpeechPath = if ($SpeechPath) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SpeechPath) } else { Join-Path $artifactBaseDir 'speech.md' }
if (-not (Test-Path -LiteralPath $resolvedGroundTruthPath)) {
    if ($Strict -or $SpecPath) {
        $styleErrors.Add("Missing ground-truth.md artifact: $resolvedGroundTruthPath") | Out-Null
    } else {
        $warnings.Add("Missing ground-truth.md artifact: $resolvedGroundTruthPath") | Out-Null
    }
}
if (-not (Test-Path -LiteralPath $resolvedSpeechPath)) {
    if ($Strict -or $SpecPath) {
        $styleErrors.Add("Missing speech.md artifact: $resolvedSpeechPath") | Out-Null
    } else {
        $warnings.Add("Missing speech.md artifact: $resolvedSpeechPath") | Out-Null
    }
}

$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Open($resolvedDeck, 0, 0, -1)

    if ($presentation.Slides.Count -lt 1) {
        $errors.Add('Deck has no slides.')
    } else {
        $firstLayout = $presentation.Slides.Item(1).CustomLayout.Name
        if ($firstLayout -ne 'Title Slide') {
            $errors.Add("Slide 1 must use official Title Slide layout, found '$firstLayout'.")
        }
        if ($spec -ne $null -and $spec.subtitle -and $spec.authorLine) {
            foreach ($shape in $presentation.Slides.Item(1).Shapes) {
                $isSubtitle = $false
                try { if ([int]$shape.PlaceholderFormat.Type -eq 4) { $isSubtitle = $true } } catch {}
                try { if ($shape.Name -match 'Subtitle') { $isSubtitle = $true } } catch {}
                if ($isSubtitle) {
                    $subtitleShapeText = ''
                    try { $subtitleShapeText = [string]$shape.TextFrame.TextRange.Text } catch {}
                    if ($subtitleShapeText -match [regex]::Escape([string]$spec.authorLine)) {
                        $styleErrors.Add('Slide 1 subtitle placeholder contains authorLine; keep deck subtitle and presenter/organization line separate.') | Out-Null
                    }
                }
            }
        }
    }

    $timingSlides = 0
    $totalPictures = 0
    $totalGroups = 0
    $totalTables = 0
    $codeLikeTextSlides = 0
    $artifactSlides = 0
    $evidenceSlides = 0
    $sourceDerivedArtifactSlides = 0
    $genericHorizontalProcessSlides = 0
    $sparseSectionDividerCount = 0
    $emptyTextShapeCount = 0
    $deadTableSlideCount = 0
    $totalShapeCount = 0
    $notesSlides = 0
    $sectionTitles = New-Object System.Collections.Generic.List[string]
    $slideTexts = @{}
    $patternCounts = @{}
    foreach ($slide in $presentation.Slides) {
        $layout = $slide.CustomLayout.Name
        $totalShapeCount += [int]$slide.Shapes.Count
        if ($allowedLayouts -notcontains $layout) {
            $errors.Add("Slide $($slide.SlideIndex) uses non-template layout '$layout'.")
        }
        if ($layout -eq 'Blank') {
            $errors.Add("Slide $($slide.SlideIndex) uses Blank layout; use official content layouts unless the user explicitly asks for a blank canvas.")
        }
        if ($slide.TimeLine.MainSequence.Count -gt 0) {
            $timingSlides++
        }
        if ($layout -eq 'Section Header') {
            try {
                $sectionTitle = $slide.Shapes.Title.TextFrame.TextRange.Text.Trim()
                if ($sectionTitle) { $sectionTitles.Add($sectionTitle) | Out-Null }
            } catch {}
        }

        try {
            $notesText = ''
            foreach ($noteShape in $slide.NotesPage.Shapes) {
                try {
                    if ($noteShape.HasTextFrame -and $noteShape.TextFrame.HasText) {
                        $notesText += ' ' + $noteShape.TextFrame.TextRange.Text
                    }
                } catch {}
            }
            if ($notesText -match '(?i)Reason:|Talk track:|Builds:') {
                $notesSlides++
            }
        } catch {}

        try {
            if ($slide.SlideShowTransition.EntryEffect -ne 0) {
                $warnings.Add("Slide $($slide.SlideIndex) has a slide transition; Li MathWorks decks should usually use none.")
            }
        } catch {}

        $slideHasPicture = $false
        $slideHasGroup = $false
        $slideHasTable = $false
        $slidePatterns = New-Object System.Collections.Generic.List[string]
        $nonPlaceholderShapeCount = 0
        foreach ($shape in $slide.Shapes) {
            $isPlaceholder = $false
            try {
                [void]$shape.PlaceholderFormat.Type
                $isPlaceholder = $true
            } catch {}
            if (-not $isPlaceholder) { $nonPlaceholderShapeCount++ }

            $name = ''
            try { $name = $shape.Name } catch {}

            $pattern = Get-LiPattern $shape
            if ($pattern) {
                $slidePatterns.Add($pattern) | Out-Null
                if (-not $patternCounts.ContainsKey($pattern)) { $patternCounts[$pattern] = 0 }
                $patternCounts[$pattern]++
            }

            try {
                if ([int]$shape.Type -in @(11, 13)) {
                    $totalPictures++
                    $slideHasPicture = $true
                }
                if ([int]$shape.Type -eq 6) {
                    $totalGroups++
                    $slideHasGroup = $true
                }
                if ($shape.HasTable) {
                    $totalTables++
                    $slideHasTable = $true
                }
            } catch {}

            $hasText = $false
            $hasTextFrame = $false
            try { $hasTextFrame = [bool]$shape.HasTextFrame } catch {}
            try { $hasText = ($shape.HasTextFrame -and $shape.TextFrame.HasText) } catch {}
            if ($hasTextFrame -and -not $hasText) {
                $emptyPlaceholderType = $null
                try { $emptyPlaceholderType = [int]$shape.PlaceholderFormat.Type } catch {}
                $isTextBoxShape = $false
                try { $isTextBoxShape = ([int]$shape.Type -eq 17) } catch {}
                if ($isTextBoxShape -or $emptyPlaceholderType -in @(2, 4, 7)) {
                    $emptyTextShapeCount++
                    if ($Strict) {
                        $styleErrors.Add("Slide $($slide.SlideIndex) has an empty text box or placeholder '$name' visible in edit mode.") | Out-Null
                    } else {
                        $warnings.Add("Slide $($slide.SlideIndex) has an empty text box or placeholder '$name' visible in edit mode.") | Out-Null
                    }
                }
            }
            if ($hasText -and -not $isPlaceholder) {
                try {
                    $size = [double]$shape.TextFrame.TextRange.Font.Size
                    if ($size -gt 0 -and $size -lt 12 -and $layout -notin @('Title Slide')) {
                        $warnings.Add("Slide $($slide.SlideIndex) has generated small text below 12 pt in '$name'.")
                    }
                } catch {}
            }

            if (-not $isPlaceholder) {
                try {
                    if ($shape.AutoShapeType -eq 1 -and -not $pattern) {
                        $warnings.Add("Slide $($slide.SlideIndex) uses a plain rectangle '$name'; prefer rounded rectangles, connectors, chevrons, or placeholders.")
                    }
                } catch {}

                try {
                    $overlapsLogo = ($shape.Left + $shape.Width -gt 800 -and $shape.Top -lt 70)
                    if ($overlapsLogo) {
                        $errors.Add("Slide $($slide.SlideIndex) has generated shape '$name' overlapping the template logo area.")
                    }
                } catch {}

                try {
                    $isBottomStripe = ($shape.Top -gt 500 -and $shape.Width -gt 650 -and $shape.Height -le 30)
                    if ($isBottomStripe) {
                        $errors.Add("Slide $($slide.SlideIndex) appears to have a custom bottom stripe '$name'; use template decoration only.")
                    }
                } catch {}
            }
        }

        if ($layout -eq 'Section Header' -and $UseCase -eq 'training-workshop') {
            $hasLocator = @($slidePatterns | Where-Object { $_ -match 'SectionLocator|ProgressRail' }).Count -gt 0
            $hasOutcome = @($slidePatterns | Where-Object { $_ -match 'SectionOutcome' }).Count -gt 0
            if (-not $hasLocator -or -not $hasOutcome) {
                $sparseSectionDividerCount++
                if ($Strict) {
                    $styleErrors.Add("Slide $($slide.SlideIndex) is a sparse training/workshop section divider; include outcome plus locator or next-artifact cue.") | Out-Null
                } else {
                    $warnings.Add("Slide $($slide.SlideIndex) is a sparse training/workshop section divider; include outcome plus locator or next-artifact cue.") | Out-Null
                }
            }
        }

        $slideText = ''
        foreach ($shape in $slide.Shapes) {
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $slideText += ' ' + $shape.TextFrame.TextRange.Text
                }
            } catch {}
        }
        if ($slideText -match 'Delete this slide before finalizing|PowerPoint Resources|Image Resources|Color Palette') {
            $errors.Add("Slide $($slide.SlideIndex) still contains template resource/instruction content.")
        }
        $slideTexts[[int]$slide.SlideIndex] = $slideText
        if ($slideText -match '(?i)use progressive reveal|speaker note|presenter note|build note|click to|animate this|visual check|\b(the workshop|the close|the deck|the slide)\s+should\b') {
            $warnings.Add("Slide $($slide.SlideIndex) appears to contain visible build/presenter instruction text.")
        }
        if ($slideHasTable -and $slideText -notmatch '(?i)\b(status|pass|passed|fail|risk|gap|validated|complete|pending|next|decision|recommend|evidence|result)\b') {
            $deadTableSlideCount++
            if ($Strict) {
                $styleErrors.Add("Slide $($slide.SlideIndex) has a table without visible status/outcome/risk/evidence language; make the table tell the story.") | Out-Null
            } else {
                $warnings.Add("Slide $($slide.SlideIndex) has a table without visible status/outcome/risk/evidence language; make the table tell the story.") | Out-Null
            }
        }
        $slideHasCodeLike = $false
        if ($slideText -match '(?i)\b(void|typedef|struct|#include|if\s*\(|for\s*\(|while\s*\(|Simulink|System Composer|RTDB|global variable|harness|requirement|verification|test result)\b') {
            $codeLikeTextSlides++
            $slideHasCodeLike = $true
        }
        if ($slideHasPicture -or $slideHasGroup -or $slideHasTable -or $slideText -match '(?i)\b(code|architecture|data|model|component|interface|wrapper|harness|requirement|verification|test result|RTDB|global variable|Simulink)\b') {
            $artifactSlides++
        }
        if ($slideHasPicture -or $slideHasTable -or $slideHasCodeLike -or @($slidePatterns | Where-Object { $_ -match 'ImageEvidence|ScreenshotEvidence|ScreenshotCallout|ModelScreenshot|Comparison|Results|ArtifactMap|ConceptArtifact|VModelToolMap|ProcessStateDiagram|Code' }).Count -gt 0) {
            $evidenceSlides++
        }
        if (@($slidePatterns | Where-Object { $_ -match 'ImageEvidence|ScreenshotEvidence|ScreenshotCallout|ModelScreenshot|Comparison|Results|ArtifactMap|ArtifactReview|ConceptArtifact|VModelToolMap|ProcessStateDiagram|Code' }).Count -gt 0) {
            $sourceDerivedArtifactSlides++
        }
        if (@($slidePatterns | Where-Object { $_ -match 'Process_horizontal-cards' }).Count -gt 0) {
            $genericHorizontalProcessSlides++
        }

        if ($layout -eq 'Two Content') {
            $contentTexts = @()
            foreach ($shape in $slide.Shapes) {
                try {
                    $placeholderType = [int]$shape.PlaceholderFormat.Type
                    if ($placeholderType -in @(2, 7) -and $shape.HasTextFrame) {
                        $contentTexts += $shape.TextFrame.TextRange.Text.Trim()
                    }
                } catch {}
            }
            if ($contentTexts.Count -ge 2 -and (($contentTexts[0].Length -eq 0) -or ($contentTexts[1].Length -eq 0))) {
                $warnings.Add("Slide $($slide.SlideIndex) uses Two Content layout with an empty column.")
            }
        }
    }

    if ($UseCase -eq 'training-workshop' -and $timingSlides -eq 0) {
        $warnings.Add('Training/workshop deck has no object animation timing. Add purposeful progressive reveal where it teaches the flow.')
    }
    if ($UseCase -eq 'training-workshop') {
        if ($totalPictures -eq 0 -and $totalGroups -eq 0 -and $totalTables -eq 0 -and $codeLikeTextSlides -eq 0) {
            $warnings.Add('Training/workshop deck has no pictures, groups, tables, or code-like artifact text. Add concrete technical artifacts, not only outline slides.')
        }
        if ($totalGroups -eq 0) {
            $warnings.Add('Training/workshop deck has no grouped objects. Li-style workshop diagrams should use grouped regions for process, data, architecture, and review artifacts.')
        }
        if ($artifactSlides -lt [Math]::Max(2, [Math]::Floor($presentation.Slides.Count / 4))) {
            $warnings.Add("Training/workshop deck appears light on technical artifact slides ($artifactSlides detected). Add code/data/model/test artifact slides.")
        }
        $averageShapeCount = if ($presentation.Slides.Count -gt 0) { [Math]::Round($totalShapeCount / $presentation.Slides.Count, 1) } else { 0 }
        if ($averageShapeCount -lt 3.5) {
            $warnings.Add("Average shape count is low ($averageShapeCount). Workshop decks often need richer diagrams, tables, screenshots, or grouped annotations.")
        }
        if ($notesSlides -eq 0) {
            $warnings.Add('Generated training/workshop deck has no speaker notes with reason/talk track/builds.')
        }
        if ($sectionTitles.Count -gt 0) {
            $agendaText = ''
            if ($spec -ne $null) {
                $specSlideIndex = 0
                foreach ($slideSpec in @($spec.slides)) {
                    $specSlideIndex++
                    if ([string]$slideSpec.type -in @('agenda', 'learning-path')) {
                        $deckSlideIndex = $specSlideIndex + 1
                        if ($slideTexts.ContainsKey($deckSlideIndex)) { $agendaText = $slideTexts[$deckSlideIndex] }
                        break
                    }
                }
            }
            if (-not $agendaText -and $slideTexts.ContainsKey(2)) { $agendaText = $slideTexts[2] }
            foreach ($sectionTitle in $sectionTitles) {
                if ($agendaText -and $agendaText -notmatch [regex]::Escape($sectionTitle)) {
                    $warnings.Add("Agenda slide does not mention section '$sectionTitle'.")
                }
            }
        }
        if ($sourceDerivedArtifactSlides -lt [Math]::Max(2, [Math]::Floor($presentation.Slides.Count / 5))) {
            $styleErrors.Add("Training/workshop deck has only $sourceDerivedArtifactSlides source-derived artifact/evidence slides. Use code, model screenshots, tables, artifact maps, or result evidence in each major chapter.") | Out-Null
        }
        if ($genericHorizontalProcessSlides -gt [Math]::Max(2, [Math]::Floor($presentation.Slides.Count / 4))) {
            $warnings.Add("Training/workshop deck has $genericHorizontalProcessSlides generic horizontal process-card slides. Mix in progress sidebars, screenshots, tables, demos, recap, and artifact-review patterns.") | Out-Null
        }
        if ($patternCounts.Keys.Count -lt 3) {
            $warnings.Add("Pattern diversity is low ($($patternCounts.Keys.Count) generated Li patterns). Use a mix of learning path, what-to-expect, progress, evidence, recap, demo, and artifact patterns.") | Out-Null
        }
    }

    if ($UseCase -eq 'customer-wrap-up') {
        if ($evidenceSlides -lt [Math]::Max(2, [Math]::Floor($presentation.Slides.Count / 4))) {
            $warnings.Add("Customer wrap-up/deep-dive deck has only $evidenceSlides evidence slides. Prefer project statement, implementation evidence, screenshots, tables, results, and next steps over abstract process cards.") | Out-Null
        }
        if ($timingSlides -gt [Math]::Max(2, [Math]::Floor($presentation.Slides.Count / 3))) {
            $warnings.Add("Customer wrap-up/deep-dive deck has many animated slides ($timingSlides). Keep customer evidence decks mostly static unless the build clarifies a workflow.") | Out-Null
        }
        if ($patternCounts.ContainsKey('WhatToExpect')) {
            $warnings.Add('Customer wrap-up/deep-dive deck uses workshop what-to-expect/color-band pattern. Use it only if the customer deck is actually a workshop.') | Out-Null
        }
    }

    if ($UseCase -eq 'internal-sharing') {
        $progressPatternCount = 0
        foreach ($key in $patternCounts.Keys) {
            if ($key -match 'Progress|Process|State') { $progressPatternCount += $patternCounts[$key] }
        }
        if ($progressPatternCount -eq 0) {
            $warnings.Add('Internal sharing deck has no progress/process scaffold patterns. Use workflow maps, status rails, tradeoff tables, or repeated build-state slides where appropriate.') | Out-Null
        }
    }

    $previewCount = $null
    if ($PreviewDir) {
        if (-not (Test-Path -LiteralPath $PreviewDir)) {
            $warnings.Add("Preview directory not found: $PreviewDir")
            $previewCount = 0
        } else {
            $previewCount = @(Get-ChildItem -LiteralPath $PreviewDir -Filter 'slide-*.png' -File).Count
            if ($previewCount -lt $presentation.Slides.Count) {
                $warnings.Add("Preview coverage incomplete: $previewCount previews for $($presentation.Slides.Count) slides.")
            }
        }
    }

    if ($TemplateKind -eq 'confidential') {
        $foundConfidential = $false
        foreach ($slide in $presentation.Slides) {
            foreach ($shape in $slide.Shapes) {
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText -and $shape.TextFrame.TextRange.Text -match '(?i)confidential') {
                        $foundConfidential = $true
                    }
                } catch {}
            }
        }
        if (-not $foundConfidential) {
            $errors.Add('Confidential template requested, but no visible confidential text was found.')
        }
    }

    $result = [pscustomobject]@{
        DeckPath = $resolvedDeck
        SlideCount = [int]$presentation.Slides.Count
        TimingSlideCount = $timingSlides
        TotalPictures = $totalPictures
        TotalGroups = $totalGroups
        TotalTables = $totalTables
        CodeLikeTextSlides = $codeLikeTextSlides
        ArtifactSlideCount = $artifactSlides
        EvidenceSlideCount = $evidenceSlides
        SourceDerivedArtifactSlideCount = $sourceDerivedArtifactSlides
        GenericHorizontalProcessSlideCount = $genericHorizontalProcessSlides
        SparseSectionDividerCount = $sparseSectionDividerCount
        EmptyTextShapeCount = $emptyTextShapeCount
        DeadTableSlideCount = $deadTableSlideCount
        PatternDiversityCount = $patternCounts.Keys.Count
        PatternCounts = $patternCounts
        NotesSlideCount = $notesSlides
        PreviewCount = $previewCount
        Errors = $errors
        StyleErrors = $styleErrors
        Warnings = $warnings
        Passed = ($errors.Count -eq 0 -and $styleErrors.Count -eq 0)
    }

    if ($AsJson) {
        $result | ConvertTo-Json -Depth 6
    } else {
        $result | Format-List
        if ($errors.Count -gt 0 -or $styleErrors.Count -gt 0) { exit 1 }
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
