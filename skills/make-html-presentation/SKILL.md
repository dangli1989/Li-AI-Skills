---
name: make-html-presentation
description: Build or improve animated HTML presentations for technical demos, AI/Codex workflow talks, customer project walkthroughs, or engineering enablement sessions. Use when the user asks to create a polished presentation, HTML deck, animated slide deck, presentation dashboard, slide flow, speaker-ready demo deck, or when they want MathWorks-like visual style, jelly/spring animations, visual slide review, slide reordering, or reusable presentation templates.
---

# Make HTML Presentation

## Goal

Create speaker-ready HTML presentations with a strong story, bold technical visuals, readable slides, and smooth spring/jelly motion. Treat the presenter as the owner of the message: plan the narrative first, then build visuals.

## Default Workflow

1. Clarify or infer the audience, time limit, core message, demo assets, and desired call to action.
2. Draft the story flow before editing slides. Prefer: hook, agenda, problem/risk, demo, method, guardrails, practical habits, close.
3. Start from `assets/template.html` unless the user provides an existing deck.
4. Use `assets/styles.css` and `assets/jelly.js` rather than rewriting the visual system.
5. Keep slides visual and presenter-led. Avoid dense explanatory text when a large image, diagram, or sequence can carry the point.
6. Use local assets when possible. If downloading images, save them beside the deck so the presentation works offline.
7. Add interactive controls only when they are reliable. For local-app launch buttons, include a fallback command if OS protocol registration may fail.
8. Render screenshots of important slides and inspect them visually before finalizing.
9. Validate slide count, counters, missing assets, and JavaScript syntax.

## Style Rules

Use the style in `references/style-guide.md` for MathWorks-like technical decks:

- light technical background, soft grid, blue/orange/yellow accents
- large typography and high contrast
- bold cards, wide spacing, minimal bullets
- smooth physical motion using the spring animation engine
- one clear message per slide

Do not create generic purple SaaS/AI-looking slides unless the user explicitly asks for that style.

## Slide Construction Rules

- Title slides should establish ownership and message, not just decorate.
- Agenda slides should match actual slide order after any rearrangement.
- Demo slides should show entry points, key artifacts, and fallback commands.
- Review/process slides should explain what the audience should do differently.
- Shortcut slides should prioritize editing/navigation efficiency if the talk is about CLI usage.
- Closing slides should reinforce verification and engineering ownership.

## Resources

- Use `assets/template.html` as the starting deck.
- Use `assets/styles.css` for the visual system.
- Use `assets/jelly.js` for physical spring/jelly animation.
- Use `assets/components.html` for reusable slide snippets.
- Read `references/storytelling.md` when arranging or rearranging a talk.
- Read `references/style-guide.md` when changing visual style.
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
