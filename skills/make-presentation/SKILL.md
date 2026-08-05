---
name: make-presentation
description: Build or improve polished presentations in HTML or PPTX for technical demos, AI/Codex workflow talks, customer project walkthroughs, engineering enablement sessions, training workshops, internal project sharing, or customer wrap-up reviews. Use when the user asks to create a presentation, slide deck, PPTX, HTML deck, animated deck, presentation dashboard, speaker-ready story flow, MathWorks-style deck, Li's MathWorks Presentation, visual slide review, slide reordering, or reusable presentation templates.
---

# Make Presentation

## Goal

Create speaker-ready presentations with a strong story, technical visuals, readable slides, and appropriate motion. Support both HTML and PPTX output. Treat the presenter as the owner of the message: plan the narrative first, then build visuals.

## Non-Negotiables

If every other rule in this skill is forgotten, these still hold. When rules appear to conflict, these win:

1. **No invented content.** Every fact, number, name, result, and claim on a slide must come from the user, the project materials, or `ground-truth.md`. Missing information becomes a visible placeholder plus a question to the user — never plausible filler.
2. **Content contract before slides.** Collect real source facts, assets, audience, and goal into `ground-truth.md` BEFORE outlining, and outline before building. A deck built before its facts exist is a defect.
3. **The template supplies all form for template-defined elements.** Cover page, titles, subtitles, footers, page numbers, confidential markings, and section headers get content only — never new styling, positioning, fonts, or replacement shapes.
4. **Template precedence ladder.** Build every visual element at the highest available level: (1) existing template placeholder on an official layout → (2) a documented preset from `references/li-mathworks-slide-presets.md` → (3) custom shapes, only when neither exists, using only the documented palette and fonts, with the justification recorded in `ground-truth.md`.
5. **Every slide must earn its place in the story.** Each slide answers a question the previous slide raised, and its title states its actual message. A slide that does neither is cut or moved.
6. **Nothing is done until reviewed.** Export previews of every slide, inspect them visually, and run the review checklist (sense first, then mechanics) before claiming completion.
7. **It must not read as machine-made.** Vary density, counts, and layout so no two adjacent slides look alike and the important slides are visibly the substantial ones. Vary form freely; never invent facts to do it. See `references/human-signature.md`.

## Default Workflow

1. Clarify or infer the audience, time limit, core message, demo assets, and desired call to action.
2. If the user did not specify output format, ask whether they want PPTX or HTML before building.
3. If the user asks for Li's MathWorks Presentation, ask whether the deck is confidential before building.
4. **Content contract phase.** Create `ground-truth.md` first and fill it with real inputs before any outline or slide work:
   - purpose, audience, use case, time limit, and call to action;
   - source facts: the actual project facts, results, numbers, names, and decisions the deck will present, each traceable to where it came from (user statement, file, screenshot, prior deck);
   - asset inventory: available screenshots, models, code excerpts, diagrams, tables;
   - missing inputs: facts or assets the deck needs but does not have — these become questions to the user and visible placeholders, never invented content.
5. Draft the story flow using `references/storytelling.md` for the matching use case. Every slide in the outline gets a one-line message and the ground-truth facts it will carry.
6. For normal generation, show the proposed outline in chat and wait for approval before creating the deck. If the user explicitly asks for fast/direct generation, proceed without waiting, but still create the planning artifacts.
7. Maintain the two markdown files for every generated presentation, including fast/direct requests:
   - `ground-truth.md`: purpose, audience, use case, source facts, agreed or working outline, slide intent, asset inventory, missing inputs/placeholders, custom-shape justifications, progress, decisions, and final output path.
   - `speech.md`: slide-by-slide talk track mapped to the final slide order.
8. Build content first and apply style last. Decide the story, facts, assets, slide intent, and speech before selecting visual presets, template-specific styling, or animations.
9. Choose the deck style:
   - Default HTML style: use `assets/template.html`, `assets/styles.css`, and `assets/jelly.js`.
   - Li's MathWorks Presentation: read `references/li-mathworks-presentation.md` and `references/li-mathworks-slide-presets.md`; classify the use case before choosing slide patterns. For PPTX, use the official template assets through the PPTX scripts; for HTML, start from the Li's MathWorks HTML assets.
10. Keep slides visual and presenter-led. Avoid dense explanatory text when a large image, diagram, screenshot, table with emphasis, or build sequence can carry the point.
11. Use local assets when possible. If downloading images, save them beside the deck so the presentation works offline.
12. Add interactive controls only when they are reliable. For local-app launch buttons, include a fallback command if OS protocol registration may fail.
13. Render screenshots or export previews of important slides and inspect them visually before finalizing.
14. Validate story coherence and fact traceability first, then slide count, counters, missing assets, syntax, template use, visual readability, `ground-truth.md`, and `speech.md`. Use `references/review-checklist.md`.

## Platform Notes

- The `scripts/*.ps1` tooling requires **Windows with desktop PowerPoint installed** (PowerShell + COM automation). It does not run on macOS or Linux.
- On macOS, use the Python fallback: `scripts/new_li_mathworks_pptx_mac.py` with the repo venv (`.venv/bin/python`, python-pptx). It builds from the same official templates with the same layout/placeholder rules and writes the same `ground-truth.md`/`speech.md` artifacts.
- On macOS, render previews with `scripts/render_pptx_preview.py DECK.pptx OUTDIR` (no PowerPoint or LibreOffice needed). It draws the master, layout, and slide shapes directly. It is faithful for geometry, colour, hierarchy, overlap, and density — but it approximates line breaking and effects, so final polish still needs one check in real PowerPoint.
- The Python generator writes its ledger into `ground-truth.md` between `<!-- generated:slide-ledger -->` markers and preserves everything you wrote outside them. Keep hand-authored context (purpose, assumptions, decisions, open questions) outside the markers.
- Animations cannot be authored on macOS (python-pptx limitation). Keep `builds` in the spec anyway: they land in the speaker notes and in `speech.md` as the intended reveal order, and can be applied later in PowerPoint.
- Never skip preview inspection because the primary tool is unavailable; use the fallback chain instead.

## Style Rules

Use the style in `references/style-guide.md` for the default HTML technical deck:

- light technical background, soft grid, blue/orange/yellow accents
- large typography and high contrast
- bold cards, wide spacing, minimal bullets
- smooth physical motion using the spring animation engine
- one clear message per slide

Use `references/li-mathworks-presentation.md` only when the user asks for Li's MathWorks Presentation or a MathWorks-style PPTX/HTML deck that should match Li's prior customer, internal, workshop, or MBSE training decks.

Do not create generic purple SaaS/AI-looking slides unless the user explicitly asks for that style.

## Slide Construction Rules

- Title slides should establish ownership and message, not just decorate.
- Agenda slides should match actual slide order after any rearrangement.
- Demo slides should show entry points, key artifacts, and fallback commands.
- Review/process slides should explain what the audience should do differently.
- Shortcut slides should prioritize editing/navigation efficiency if the talk is about CLI usage.
- Closing slides should reinforce verification and engineering ownership.

## Resources

- For default HTML, use `assets/template.html` as the starting deck.
- For default HTML, use `assets/styles.css` for the visual system.
- For default HTML, use `assets/jelly.js` for physical spring/jelly animation.
- For default HTML, use `assets/components.html` for reusable slide snippets.
- For Li's MathWorks PPTX, use `assets/li-mathworks-presentation/pptx/public.pptx` or `assets/li-mathworks-presentation/pptx/confidential.pptx`.
- For Li's MathWorks HTML, use `assets/li-mathworks-presentation/html/template.html` and `assets/li-mathworks-presentation/html/styles.css`.
- For Li's MathWorks PPTX on Windows, run `scripts/inspect_li_mathworks_pptx.ps1` on the chosen template before creating slides.
- For Li's MathWorks PPTX on Windows, prefer `scripts/new_li_mathworks_pptx.ps1` with a JSON slide spec instead of hand-building slides.
- For Li's MathWorks PPTX on Windows, run `scripts/review_li_mathworks_pptx.ps1` and `scripts/export_pptx_previews.ps1` before finalizing.
- For Li's MathWorks PPTX on macOS, use `scripts/new_li_mathworks_pptx_mac.py` with the same JSON slide spec format.
- See `references/example-spec.md` for a complete worked example: JSON spec, `ground-truth.md`, and `speech.md` for one short deck.
- Read `references/human-signature.md` before finalising any deck: it is what keeps the output from reading as machine-generated.
- Read `references/storytelling.md` when planning, arranging, or rearranging a talk.
- Read `references/style-guide.md` when changing visual style.
- Read `references/li-mathworks-presentation.md` for Li's MathWorks Presentation.
- Read `references/li-mathworks-slide-presets.md` when creating Li's MathWorks PPTX slides.
- Read `references/review-checklist.md` before final validation.

## Validation

On Windows for HTML decks:

```powershell
.\scripts\validate_deck.ps1 -DeckPath <path-to-html>
.\scripts\render_slide.ps1 -DeckPath <path-to-html> -SlideIndex 2 -OutPath <path-to-png>
```

On macOS for PPTX decks, export previews via the Python helper or LibreOffice, then inspect.

Inspect rendered screenshots for overlap, clipped text, unreadable labels, weak contrast, and whether the slide communicates the intended message without presenter explanation. Then run the sense checks in `references/review-checklist.md`: does each slide trace to `ground-truth.md`, does each slide follow from the previous one, and does the deck answer the audience's question.
