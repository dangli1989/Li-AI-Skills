# Li's MathWorks Slide Presets

Use these presets with `scripts/new_li_mathworks_pptx.ps1` for Li's MathWorks PPTX decks. The presets are a style layer on top of the official MathWorks template, not a replacement for it.

## Global Rules

- Start from `public.pptx` or `confidential.pptx`.
- Use template layouts and placeholders first.
- Create and maintain `ground-truth.md` and `speech.md` for every deck, including fast/direct generation.
- Build the story, outline, slide intent, evidence plan, and talk track before applying these presets.
- Run PowerPoint automation scripts serially; do not parallelize generation, review, or preview export.
- Do not create custom cover pages, footers, bottom bars, logos, section backgrounds, or title bands.
- Use larger text and fewer words than a dense report slide.
- Use low-saturation color coding and grouped visuals.
- Avoid plain sharp square boxes. Prefer rounded rectangles, chevrons, callouts, and connector-based flows.
- Use transparent/no fill when a shape should blend with the background.
- Use animation to teach flow, not to decorate.
- Prefer diagrams, screenshots, staged builds, evidence artifacts, model/code excerpts, emphasized tables, and callouts over text-only slides. If a text-only slide is necessary, record the reason in `ground-truth.md`.
- For training/workshop decks, include at least one concrete technical artifact per major chapter: source code excerpt, data ownership map, architecture/model view, wrapper/harness, verification matrix, screenshot, or result evidence.
- Put build guidance in speaker notes, not visible slide text.
- Add `reason`, `talkTrack`, and optional `builds` fields to every JSON slide spec so the deck records why each slide exists and how it should be presented.
- Add `assetStatus`, `placeholderStatus`, `progressLabel`, and `missingInputs` fields when they apply. Keep those fields synchronized with `ground-truth.md`.
- Build the deck from the full available Li example corpus. Do not infer the style from one deck or one generated draft.
- Choose or combine presets by use case. Training decks need teaching page types, customer wrap-ups need evidence pages, and internal sharing decks need progress/review scaffolds.
- Prefer a single source of truth for progress labels. Use top-level `progressLabels` plus per-slide `activeProgressLabel` when possible instead of retyping locator arrays.
- Do not add a color-code declaration slide unless the same color code is used consistently through the deck.
- Do not keep unused or empty template/generator text boxes. Delete empty generated boxes and clear unused placeholders before delivery.

## Presets

`agenda`

- Layout: prefer `Agenda` only when the template exposes editable title/content placeholders cleanly. In the V26 template, use `Title and Content` for agenda generation if the `Agenda` layout shows master prompt text in exported previews.
- Use title and content placeholders.
- Use 4-6 agenda items maximum.
- Use primary item text around 18-22 pt and supporting text around 14-16 pt.
- Avoid separate number boxes unless they already exist in the template or source style.

`section`

- Layout: `Section Header`.
- Use the title placeholder.
- For training/workshop decks, include a short outcome and optional mini process locator.
- No custom solid-blue divider slides.
- Do not overuse sparse section slides. For workshop delivery, follow a section divider with a concrete artifact, demo, exercise, or progress-review page.
- Generate locator labels from the shared progress list when possible. Duplicate labels in one locator are a defect.

`what-to-expect`

- Layout: `Title and Content`.
- Use this early in training/workshop decks.
- Show workshop format, audience action, color/page-type coding, demo/exercise expectations, and discussion expectations.
- Use a large framed content area with colored bands or subpanels. Keep visible wording concise; leave detailed facilitation notes in speaker notes.
- Match the MBSE training pattern: one clear outer frame and page-type colors for best practice, why, how, what to do, demo/exercise, and discussion.

`chapter-divider`

- Layout: `Section Header` or `Title and Content` depending on whether the chapter needs a sparse reset or a progress scaffold.
- Use official title placeholders. Do not create custom title bands.
- Include outcome and a progress locator only when it helps the audience know where they are.
- For training decks, prefer a divider that sets the next exercise/demo target rather than a decorative divider.

`progress-sidebar`

- Layout: `Title and Content`.
- Use for internal process sharing or workshop process review.
- Build a left rail with the full process list and one highlighted active step.
- Main canvas should contain an artifact, workflow diagram, screenshot, or state change. Do not leave the main area as only a title and three cards.
- Reuse the same rail across a repeated build series so audiences can track progress.

`process-state-diagram`

- Layout: `Title and Content`.
- Use for internal sharing or workshop process review when the audience needs to see what changed between repeated states.
- Keep a stable workflow scaffold and add visible state marks: X, check, clock, warning, owner, risk, or evidence labels.
- Use low-saturation process regions and semantic status colors. Avoid floating small pills that overlap larger cards.
- Reveal the scaffold first, then each state change in the talk order.

`process-build-series`

- Layout: `Title and Content`.
- Use when several slides explain the same workflow over time.
- Keep the scaffold stable and add visible build states: icons, X marks, checks, clocks, arrows, status labels, or evidence screenshots.
- Each repeated slide must add information. Do not create several nearly identical horizontal process-card pages.

`process`

- Layout: `Title and Content`.
- Use the title placeholder.
- Build the process inside the content area with rounded rectangles, chevrons, or connector arrows.
- Use light-blue, pale-yellow, pale-green, light-cyan, and neutral fills.
- Reveal steps in click order for workshop decks.
- Prefer builds that reveal the first node, then each `arrow + next node` together.
- Use this for compact overviews only. If the deck needs to teach or review a real workflow, prefer `progress-sidebar` or `process-build-series`.

`decision`

- Layout: `Title and Content`.
- Use 2-4 grouped callouts in the content area.
- Keep each callout to one short phrase plus optional speaker-note-level detail.
- Use subtle fill colors and theme outlines.

`two-content`

- Layout: `Two Content`.
- Use left and right content placeholders.
- Use this for comparing current vs. target, input vs. output, or model artifact vs. validation artifact.
- Keep body text at 14 pt or larger.
- If either column is empty, do not use this preset. Use `Title and Content` or add the missing comparison side.
- For workshop delivery, prefer paired rows/tables when the presenter should compare items one by one.

`artifact-map`

- Layout: `Title and Content`.
- Use grouped rounded regions to show concrete artifacts: source code, data owner, component boundary, model artifact, harness, or result.
- Use semantic color coding: gray = current/legacy, blue = model/workflow, orange = conversion/action, green = validated/pass, red = risk/blocker, yellow = review highlight.
- Use this when a chapter would otherwise be an abstract outline.
- Create composite grouped regions when possible: icon/evidence, label, detail, role/status. Six independent cards are acceptable only for a quick summary, not for the main teaching artifact.

`concept-artifact`

- Layout: `Title and Content`.
- Use when a training/workshop or technical deck needs both the idea and the concrete artifact on the same page.
- Left side: concept, principle, or decision in 2-4 short statements.
- Right side: screenshot/model/code/table placeholder or evidence panel with 2-4 callouts.
- Use a thin low-contrast frame, transparent/no-fill background, and labels large enough to read.
- Track missing screenshots or source artifacts in `ground-truth.md` and show a visible placeholder only when the asset is not yet available.

`v-model-tool-map`

- Layout: `Title and Content`.
- Use for MBD/MBSE lifecycle explanations, System Composer training, or conversion workflow maps.
- Use native PowerPoint shapes: V/process path, low-saturation phase regions, orange-outlined MathWorks tool labels, and side evidence/benefit callouts.
- Reveal the path in order for workshops; keep mostly static for customer wrap-up unless the build clarifies the workflow.
- Do not use this only as decoration. The title and labels must state what decision or method the diagram teaches.

`code-review-excerpt`

- Layout: `Title and Content`.
- Use a readable monospace excerpt plus 2-4 callouts.
- Use this instead of copying an entire dense code slide when the teaching point is a specific pattern.
- Keep the code excerpt short enough to read; leave full code detail to the speaker or backup.

`image-evidence`

- Layout: `Title and Content` or `Two Content`.
- Use screenshots, exported model views, or diagrams as evidence.
- Recreate labels as PowerPoint text if screenshot text is too small.
- Do not rely on unreadable embedded screenshot text.
- Use for customer wrap-up, MBSE training, System Composer demos, requirements/test evidence, Simulink models, and implementation screenshots.
- Prefer one large screenshot plus 2-4 callouts, or two side-by-side screenshots with clear captions.

`comparison-table`

- Layout: `Title and Content`.
- Use when the presenter needs to compare current vs target, option A vs option B, or source artifact vs model artifact.
- Tables may use smaller text than process cards when dense technical comparison is the point. Keep headers bold and use light fills.
- Do not replace a real comparison with three generic decision cards.
- Make tables tell a story without the presenter: use status coloring, row emphasis, outcome/risk labels, callouts, or row-by-row builds. A plain grid of text is not enough for Li's style.

`results-table`

- Layout: `Title and Content`.
- Use for customer wrap-up/deep-dive and verification summary slides.
- Include result name, evidence/source, status, and next action or risk.
- Use green/red/yellow status color sparingly and consistently.
- Make the status/result visible at a glance with semantic cell fills, light row shading, and a clear next-action or risk column.

`model-screenshot`

- Layout: `Title and Content`.
- Use one main model, architecture, requirements, project, or test screenshot with callouts.
- Add labels as PowerPoint text when screenshot labels are too small.
- Use this instead of abstract model-process cards when a concrete System Composer or Simulink artifact exists.

`demo-exercise`

- Layout: `Title and Content`.
- Use for training/workshop decks.
- Show task, starting artifact, expected attendee action, and output artifact.
- Use cyan/light-blue demo panels and optional discussion/risk panel. Put facilitation detail in notes.

`recap`

- Layout: `Title and Content`.
- Use after chapters or sessions in training decks.
- Summarize takeaways with a mix of bullets, screenshots, and small evidence panels.
- Recap slides can be denser than concept slides when they serve as memory anchors.

`team-context`

- Layout: `Title and Content`.
- Use for internal/customer kickoff or project-sharing decks when attendees need to know roles.
- Prefer real team/project context, role boxes, and contact/ownership information. Do not use named-person folders or hard-coded personal paths in reusable assets.

`imported-source-slide`

- Use only when there is a concrete reason to preserve the original slide.
- Acceptable reasons: customer-approved artifact, complex source animation, screenshot/evidence preservation, or exact legacy diagram needed for discussion.
- After import, check whether neighboring slides visually clash. If they do, recreate the idea using template layouts instead of leaving a pasted slide in the flow.
