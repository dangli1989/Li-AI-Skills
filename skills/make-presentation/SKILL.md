---
name: make-presentation
description: Build or improve polished presentations in HTML or PPTX for technical demos, AI/Codex workflow talks, customer project walkthroughs, engineering enablement sessions, training workshops, internal project sharing, or customer wrap-up reviews. Use when the user asks to create a presentation, slide deck, PPTX, HTML deck, animated deck, presentation dashboard, speaker-ready story flow, MathWorks-style deck, Li's MathWorks Presentation, visual slide review, slide reordering, or reusable presentation templates.
---

# Make Presentation

## Goal

Create speaker-ready presentations with a strong story, technical visuals, readable slides, and appropriate motion. Support both HTML and PPTX output. Treat the presenter as the owner of the message: plan the narrative first, then build visuals.

## Default Workflow

1. Clarify or infer the audience, time limit, core message, demo assets, and desired call to action.
2. If the user did not specify output format, ask whether they want PPTX or HTML before building.
3. If the user asks for Li's MathWorks Presentation, ask whether the deck is confidential before building.
4. Draft the story flow before editing slides. Prefer: hook, agenda, problem/risk, demo, method, guardrails, practical habits, close.
5. For normal generation, show the proposed outline in chat and wait for approval before creating the deck. If the user explicitly asks for fast/direct generation, proceed without waiting, but still create the planning artifacts below.
6. Create and maintain two markdown files for every generated presentation, including fast/direct requests:
   - `ground-truth.md`: purpose, audience, use case, source facts, agreed or working outline, slide intent, asset inventory, missing inputs/placeholders, progress, decisions, and final output path.
   - `speech.md`: slide-by-slide talk track mapped to the final slide order.
7. Build content first and apply style last. Decide the story, facts, assets, slide intent, and speech before selecting visual presets, template-specific styling, or animations.
8. Choose the deck style:
   - Default HTML style: use `assets/template.html`, `assets/styles.css`, and `assets/jelly.js`.
   - Li's MathWorks Presentation: read `references/li-mathworks-presentation.md` and `references/li-mathworks-slide-presets.md`; classify the use case before choosing slide patterns. This style is a use-case-aware combination of Li's prior MathWorks decks, not a clone of one example deck. For PPTX, use the official template assets through the PPTX scripts; for HTML, start from the Li's MathWorks HTML assets.
9. Keep slides visual and presenter-led. Avoid dense explanatory text when a large image, diagram, screenshot, table with emphasis, or build sequence can carry the point.
10. Use local assets when possible. If downloading images, save them beside the deck so the presentation works offline.
11. Add interactive controls only when they are reliable. For local-app launch buttons, include a fallback command if OS protocol registration may fail.
12. Render screenshots or export previews of important slides and inspect them visually before finalizing.
13. Validate slide count, counters, missing assets, syntax, template use, visual readability, `ground-truth.md`, and `speech.md`.

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
- For Li's MathWorks PPTX, run `scripts/inspect_li_mathworks_pptx.ps1` on the chosen template before creating slides.
- For Li's MathWorks PPTX, prefer `scripts/new_li_mathworks_pptx.ps1` with a JSON slide spec instead of hand-building slides.
- For Li's MathWorks PPTX, run `scripts/review_li_mathworks_pptx.ps1` and `scripts/export_pptx_previews.ps1` before finalizing.
- For Li's MathWorks PPTX, use the full local style-intake corpus when available: `generated/style-intake/li-mathworks-presentation/examples/*.pptx` and MBSE training slides under the source repo. Do not infer the style from a single example deck.
- Read `references/storytelling.md` when arranging or rearranging a talk.
- Read `references/style-guide.md` when changing visual style.
- Read `references/li-mathworks-presentation.md` for Li's MathWorks Presentation.
- Read `references/li-mathworks-slide-presets.md` when creating Li's MathWorks PPTX slides.
- Read `references/review-checklist.md` before final validation.

## Validation

Run:

```powershell
.\scripts\validate_deck.ps1 -DeckPath <path-to-html>
```

For visual inspection, render a slide:

```powershell
.\scripts\render_slide.ps1 -DeckPath <path-to-html> -SlideIndex 2 -OutPath <path-to-png>
```

Inspect rendered screenshots for overlap, clipped text, unreadable labels, weak contrast, and whether the slide communicates the intended message without presenter explanation.
